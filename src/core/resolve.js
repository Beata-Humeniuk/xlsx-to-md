'use strict';

// Finding a saved filter's data in a workbook — the same file later, or a
// different one. The sheet is found by name, then by its headers; each range
// by its Excel table name, then by its header names (columns may have moved,
// rows may have been added), and only as a last resort by its address.

const { parseRange, formatRange } = require('./address');
const { prepare, dataEnd, findHeaderRow, looksLikeHeader } = require('./region');

function savedHeaders(source) {
  const names = [];
  for (const r of source.ranges) for (const h of r.headers || []) if (h) names.push(h);
  return names;
}

function headerScore(sheet, names) {
  if (!names.length) return 0;
  const hit = findHeaderRow(sheet, names);
  return hit ? hit.found / hit.total : 0;
}

function pickSheet(workbook, source, notes) {
  const sheets = workbook.sheets;
  // A CSV file (or a one-sheet workbook) has nothing to choose from; its
  // sheet name follows the file name and is no reason for a note.
  if (sheets.length === 1) return 0;
  const names = savedHeaders(source);
  const wanted = String(source.sheet || '').toLowerCase();
  const byName = sheets.findIndex((s) => s.name.toLowerCase() === wanted);
  if (byName >= 0) {
    if (!names.length || headerScore(workbook.loadSheet(byName), names) >= 0.5) return byName;
  }
  if (names.length) {
    let best = -1;
    let bestScore = 0;
    sheets.forEach((s, i) => {
      if (i === byName) return;
      const score = headerScore(workbook.loadSheet(i), names);
      if (score > bestScore + 1e-9 || (score === bestScore && best >= 0 && sheets[best].hidden && !s.hidden)) {
        best = i;
        bestScore = score;
      }
    });
    if (best >= 0 && bestScore >= 0.5) {
      if (source.sheet) notes.push({ code: 'sheetByHeaders', sheet: source.sheet, used: sheets[best].name });
      return best;
    }
  }
  if (byName >= 0) return byName;
  const visible = sheets.findIndex((s) => !s.hidden);
  const used = visible >= 0 ? visible : 0;
  if (source.sheet) notes.push({ code: 'sheetMissing', sheet: source.sheet, used: sheets[used].name });
  return used;
}

function resolveRange(sheet, saved, headerRow, notes) {
  if (saved.table) {
    const t = (sheet.tables || []).find((x) => x.name.toLowerCase() === saved.table.toLowerCase());
    if (t) return { range: { ...t.range }, table: t.name };
  }
  const fixed = parseRange(saved.ref, sheet);
  const height = fixed ? fixed.bottom - fixed.top : 0;
  const names = (saved.headers || []).filter(Boolean);
  if (headerRow && names.length) {
    const hit = findHeaderRow(sheet, names);
    if (hit && hit.found * 2 >= hit.total) {
      const top = hit.row;
      const bottom = saved.toEnd ? dataEnd(sheet, top, hit.left, hit.right) : Math.min(sheet.rows - 1, top + height);
      const range = { top, left: hit.left, bottom: Math.max(top, bottom), right: hit.right };
      if (hit.found < hit.total) notes.push({ code: 'headersPartly', range: formatRange(range), found: hit.found, total: hit.total });
      return { range, table: null };
    }
    notes.push({ code: 'headersNotFound', ref: saved.ref });
  }
  if (!fixed) return null;
  const range = { ...fixed };
  if (saved.toEnd) range.bottom = dataEnd(sheet, range.top, range.left, range.right);
  return { range, table: null };
}

// workbook: { sheets: [{ name, hidden }], loadSheet(i) }; recipe: normalized.
// Returns { sheetIndex, ranges: [{ top, left, bottom, right }], notes }.
function resolveSource(workbook, recipe) {
  const notes = [];
  const source = recipe.source;
  const sheetIndex = pickSheet(workbook, source, notes);
  const sheet = prepare(workbook.loadSheet(sheetIndex));
  const ranges = [];
  for (const saved of source.ranges) {
    const hit = resolveRange(sheet, saved, source.headerRow, notes);
    if (hit) ranges.push(hit.range);
  }
  return { sheetIndex, ranges, notes };
}

// What a saved filter remembers about one selected range: its address, the
// Excel table it is (if it is exactly one), its header names, and whether it
// runs to the end of the data — then it keeps doing so as rows are added.
function describeRange(sheet, range, headerRow) {
  prepare(sheet);
  const table = (sheet.tables || []).find((t) => t.range.top === range.top && t.range.left === range.left
    && t.range.bottom === range.bottom && t.range.right === range.right);
  const headers = [];
  if (headerRow) {
    for (let c = range.left; c <= range.right && c < sheet.cols; c++) headers.push(sheet.text[range.top][c] || null);
  }
  const end = dataEnd(sheet, range.top, range.left, Math.min(range.right, sheet.cols - 1));
  return {
    ref: formatRange(range),
    table: table ? table.name : null,
    headers: headerRow ? headers : null,
    toEnd: range.bottom >= end,
  };
}

function headerRowGuess(sheet, ranges) {
  return ranges.length > 0 && ranges.every((r) => looksLikeHeader(sheet, r));
}

module.exports = { resolveSource, describeRange, headerRowGuess };
