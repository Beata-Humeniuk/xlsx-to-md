'use strict';

// Questions about where data sits on a sheet: the block of filled cells
// around a cell (what Ctrl+* selects in Excel), where a table ends, whether
// a range starts with a header row, and where a set of headers is now.

const { headerKey } = require('./columns');

// Adds lookup structures to a sheet from the workbook reader (once).
function prepare(sheet) {
  if (sheet.prepared) return sheet;
  const mergeOf = new Map();
  for (const m of sheet.merges || []) {
    for (let r = m.top; r <= m.bottom && r < sheet.rows; r++) {
      for (let c = m.left; c <= m.right && c < sheet.cols; c++) {
        if (r !== m.top || c !== m.left) mergeOf.set(r * 16384 + c, m);
      }
    }
  }
  Object.defineProperty(sheet, 'prepared', {
    value: {
      hiddenRows: new Set(sheet.hiddenRows || []),
      hiddenCols: new Set(sheet.hiddenCols || []),
      mergeOf,
    },
    enumerable: false,
  });
  return sheet;
}

function text(sheet, r, c) {
  return r < sheet.rows && c < sheet.cols ? sheet.text[r][c] : '';
}

// The cell as data: merged areas show the value of their top-left cell in
// every cell they cover when `fillMerged` is on.
function cell(sheet, r, c, fillMerged) {
  if (r >= sheet.rows || c >= sheet.cols || r < 0 || c < 0) return { text: '', kind: '', num: null };
  if (fillMerged) {
    const m = prepare(sheet).prepared.mergeOf.get(r * 16384 + c);
    if (m) return cell(sheet, m.top, m.left, false);
  }
  return { text: sheet.text[r][c], kind: sheet.kind[r][c], num: sheet.num[r][c] };
}

function rowEmpty(sheet, r, left, right) {
  if (r >= sheet.rows) return true;
  const row = sheet.text[r];
  for (let c = left; c <= right && c < sheet.cols; c++) if (row[c] !== '') return false;
  return true;
}

function colEmpty(sheet, c, top, bottom) {
  if (c >= sheet.cols) return true;
  for (let r = top; r <= bottom && r < sheet.rows; r++) if (sheet.text[r][c] !== '') return false;
  return true;
}

// The rectangle of filled cells connected to (row, col), grown until it is
// bordered by empty rows and columns — Excel's "current region".
function currentRegion(sheet, row, col) {
  const r = { top: row, left: col, bottom: row, right: col };
  let grew = true;
  while (grew) {
    grew = false;
    if (r.top > 0 && !rowEmpty(sheet, r.top - 1, Math.max(0, r.left - 1), r.right + 1)) { r.top--; grew = true; }
    if (r.bottom + 1 < sheet.rows && !rowEmpty(sheet, r.bottom + 1, Math.max(0, r.left - 1), r.right + 1)) { r.bottom++; grew = true; }
    if (r.left > 0 && !colEmpty(sheet, r.left - 1, Math.max(0, r.top - 1), r.bottom + 1)) { r.left--; grew = true; }
    if (r.right + 1 < sheet.cols && !colEmpty(sheet, r.right + 1, Math.max(0, r.top - 1), r.bottom + 1)) { r.right++; grew = true; }
  }
  return r;
}

// The last row of the data that starts at `top`: the row before the first
// completely empty row within the columns left..right.
function dataEnd(sheet, top, left, right) {
  let r = top;
  while (r + 1 < sheet.rows && !rowEmpty(sheet, r + 1, left, right)) r++;
  return r;
}

// The smallest range holding every filled cell; null for an empty sheet.
function usedRange(sheet) {
  let top = -1;
  let bottom = -1;
  let left = Infinity;
  let right = -1;
  for (let r = 0; r < sheet.rows; r++) {
    const row = sheet.text[r];
    for (let c = 0; c < sheet.cols; c++) {
      if (row[c] === '') continue;
      if (top < 0) top = r;
      bottom = r;
      if (c < left) left = c;
      if (c > right) right = c;
    }
  }
  return top < 0 ? null : { top, left, bottom, right };
}

// Whether the first row of the range reads as column headers: every filled
// cell is text, at least half of them are filled, and there is data below.
function looksLikeHeader(sheet, range) {
  if (range.bottom <= range.top) return false;
  let filled = 0;
  let width = 0;
  for (let c = range.left; c <= range.right; c++) {
    width++;
    const cl = cell(sheet, range.top, c, true);
    if (cl.text === '') continue;
    if (cl.kind !== 's') return false;
    filled++;
  }
  return filled > 0 && filled * 2 >= width;
}

// Looks for the row holding the given header names. Returns
// { row, left, right, found, total } for the best row (most names found,
// higher on the sheet first), or null when not one name is there.
function findHeaderRow(sheet, names, limitRows) {
  const wanted = [...new Set(names.filter(Boolean).map(headerKey))];
  if (!wanted.length) return null;
  const maxRow = Math.min(sheet.rows, limitRows || 2000);
  let best = null;
  for (let r = 0; r < maxRow; r++) {
    const row = sheet.text[r];
    const at = new Map();
    for (let c = 0; c < sheet.cols; c++) {
      if (row[c] === '') continue;
      const k = headerKey(row[c]);
      if (!at.has(k)) at.set(k, c);
    }
    let found = 0;
    let left = Infinity;
    let right = -1;
    for (const w of wanted) {
      if (!at.has(w)) continue;
      found++;
      left = Math.min(left, at.get(w));
      right = Math.max(right, at.get(w));
    }
    if (found && (!best || found > best.found)) best = { row: r, left, right, found, total: wanted.length };
    if (best && best.found === wanted.length) break;
  }
  if (!best) return null;
  // Columns added at the edges of the header row belong to the table too.
  while (best.left > 0 && text(sheet, best.row, best.left - 1) !== '') best.left--;
  while (best.right + 1 < sheet.cols && text(sheet, best.row, best.right + 1) !== '') best.right++;
  return best;
}

module.exports = { prepare, cell, rowEmpty, currentRegion, dataEnd, usedRange, looksLikeHeader, findHeaderRow };
