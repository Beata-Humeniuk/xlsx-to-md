'use strict';

// From a sheet, the selected ranges and a recipe to the tables that go into
// the Markdown: header row, visible columns, row conditions, sorting and the
// chosen columns in the chosen order.

const { cell, prepare, rowEmpty } = require('./region');
const { rangeColumns, findColumn } = require('./columns');
const { formatRange } = require('./address');
const { formatGeneral, serialToParts } = require('./numfmt');
const filter = require('./filter');

function isoOfSerial(serial, date1904) {
  if (serial == null) return '';
  const p = serialToParts(serial, date1904);
  const pad = (n) => String(n).padStart(2, '0');
  let out = p.year + '-' + pad(p.month) + '-' + pad(p.day);
  const s = Math.round(p.seconds);
  if (s) {
    out += ' ' + pad(Math.floor(s / 3600)) + ':' + pad(Math.floor((s % 3600) / 60));
    if (s % 60) out += ':' + pad(s % 60);
  }
  return out;
}

function outputText(c, options, date1904) {
  if (options.values !== 'raw' || c.num == null) return c.text;
  if (c.kind === 'n') return formatGeneral(c.num);
  if (c.kind === 'd') return isoOfSerial(c.num, date1904);
  return c.text;
}

// One selected range read into columns and data rows.
function readRange(sheet, range, recipe) {
  const o = recipe.options;
  const { hiddenRows, hiddenCols } = prepare(sheet).prepared;
  const bottom = Math.min(range.bottom, sheet.rows - 1);
  const right = Math.min(range.right, Math.max(sheet.cols - 1, range.left));
  const entries = [];
  for (let c = range.left; c <= right; c++) {
    if (o.skipHiddenColumns && hiddenCols.has(c)) continue;
    entries.push({ text: recipe.source.headerRow ? cell(sheet, range.top, c, o.fillMerged).text : '', col: c });
  }
  const columns = rangeColumns(entries, recipe.source.headerRow);
  const rows = [];
  const first = recipe.source.headerRow ? range.top + 1 : range.top;
  let hiddenSkipped = 0;
  for (let r = first; r <= bottom; r++) {
    if (o.skipHiddenRows && hiddenRows.has(r)) {
      hiddenSkipped++;
      continue;
    }
    if (o.skipEmptyRows && rowEmpty(sheet, r, range.left, right) && !anyMerged(sheet, r, columns, o)) continue;
    const cells = new Map();
    for (const col of columns) cells.set(col.id, cell(sheet, r, col.col, o.fillMerged));
    rows.push({ row: r, cells });
  }
  return { range, columns, rows, hiddenSkipped };
}

function anyMerged(sheet, r, columns, o) {
  if (!o.fillMerged) return false;
  return columns.some((col) => cell(sheet, r, col.col, true).text !== '');
}

function sortRows(rows, column, dir, date1904) {
  if (!column) return rows;
  const id = column.id;
  const sign = dir === 'desc' ? -1 : 1;
  const key = (row) => row.cells.get(id) || { text: '', kind: '', num: null };
  return rows
    .map((row, i) => ({ row, i }))
    .sort((a, b) => {
      const ca = key(a.row);
      const cb = key(b.row);
      // Empty cells go last whichever way the sort runs.
      if (ca.text === '' && cb.text !== '') return 1;
      if (cb.text === '' && ca.text !== '') return -1;
      let d;
      if (ca.num != null && cb.num != null) d = ca.num - cb.num;
      else {
        const na = filter.parseNumber(ca.text);
        const nb = filter.parseNumber(cb.text);
        d = na !== null && nb !== null ? na - nb
          : filter.fold(ca.text).localeCompare(filter.fold(cb.text), undefined, { numeric: true });
      }
      return d * sign || a.i - b.i;
    })
    .map((x) => x.row);
}

// sheet: from the workbook reader; ranges: [{ top, left, bottom, right }]
// (resolved); recipe: normalized. Returns
// {
//   available: [{ id, name, letter, type }],   every column of the selection
//   tables: [{ label, columns: [{ id, name, letter, type }], rows: [[text]], types }],
//   stats: { rows, kept, hiddenSkipped },
//   keptRows: Set of sheet rows that made it into the output,
//   excludedCols: Set of sheet columns of the selection left out of it,
//   missing: { columns: [ref], conditions: [condition] },
// }
function extract(sheet, ranges, recipe, context = {}) {
  prepare(sheet);
  const date1904 = !!sheet.date1904;
  const parts = ranges.map((r) => readRange(sheet, r, recipe));

  const available = [];
  const byId = new Map();
  for (const p of parts) {
    for (const c of p.columns) {
      if (byId.has(c.id)) continue;
      const entry = { id: c.id, name: c.name, letter: c.letter, type: 'text' };
      byId.set(c.id, entry);
      available.push(entry);
    }
  }
  for (const a of available) {
    const cells = [];
    for (const p of parts) for (const row of p.rows) if (row.cells.has(a.id)) cells.push(row.cells.get(a.id));
    a.type = filter.columnType(cells);
  }

  const missingConditions = [];
  const active = [];
  for (const cond of recipe.filter.conditions) {
    if (!cond.column) continue;
    const col = findColumn(available, cond.column);
    if (!col) {
      missingConditions.push(cond);
      continue;
    }
    if (filter.isActive(cond, cond.values)) active.push({ cond, id: col.id });
  }
  const matchAll = recipe.filter.match !== 'any';
  const passes = (row) => {
    if (!active.length) return true;
    const results = active.map(({ cond, id }) => filter.test(cond, row.cells.get(id), cond.values, date1904));
    return matchAll ? results.every(Boolean) : results.some(Boolean);
  };

  const sortCol = recipe.sort.column ? findColumn(available, recipe.sort.column) : null;

  let output;
  const missingColumns = [];
  if (recipe.columns.length) {
    output = [];
    for (const c of recipe.columns) {
      const hit = findColumn(available, c);
      if (!hit) {
        if (c.include) missingColumns.push(c);
        continue;
      }
      if (c.include && !output.includes(hit)) output.push(hit);
    }
  } else {
    output = available.slice();
  }

  // Sheet columns inside the selection that do not reach the output.
  const outputIds = new Set(output.map((c) => c.id));
  const excludedCols = new Set();
  for (const p of parts) {
    const inPart = new Set(p.columns.map((c) => c.col));
    for (let c = p.range.left; c <= p.range.right; c++) if (!inPart.has(c)) excludedCols.add(c);
    for (const c of p.columns) if (!outputIds.has(c.id)) excludedCols.add(c.col);
  }

  let total = 0;
  let kept = 0;
  let hiddenSkipped = 0;
  const keptRows = new Set();
  const filtered = parts.map((p) => {
    total += p.rows.length;
    hiddenSkipped += p.hiddenSkipped;
    const rows = sortRows(p.rows.filter(passes), sortCol, recipe.sort.dir, date1904);
    kept += rows.length;
    for (const row of rows) keptRows.add(row.row);
    return { ...p, rows };
  });

  const sameHeaders = filtered.every((p) => p.columns.map((c) => c.id).join('\n') === filtered[0].columns.map((c) => c.id).join('\n'));
  const merge = filtered.length > 1 && (recipe.options.multiRange === 'merge' || (recipe.options.multiRange === 'auto' && sameHeaders));

  const toTable = (label, pieces, columns) => {
    let rows = [];
    for (const p of pieces) rows = rows.concat(p.rows);
    if (merge && sortCol) rows = sortRows(rows, sortCol, recipe.sort.dir, date1904);
    let cols = columns;
    if (recipe.options.dropEmptyColumns) {
      cols = cols.filter((c) => rows.some((row) => {
        const v = row.cells.get(c.id);
        return v && v.text !== '';
      }));
    }
    return {
      label,
      columns: cols,
      types: cols.map((c) => c.type),
      rows: rows.map((row) => cols.map((c) => {
        const v = row.cells.get(c.id);
        return v ? outputText(v, recipe.options, date1904) : '';
      })),
    };
  };

  let tables;
  if (!filtered.length) tables = [];
  else if (merge || filtered.length === 1) {
    tables = [toTable(filtered.map((p) => labelOf(p.range, context)).join(', '), filtered, output)];
  } else {
    tables = filtered.map((p) => {
      const ids = new Set(p.columns.map((c) => c.id));
      return toTable(labelOf(p.range, context), [p], output.filter((c) => ids.has(c.id)));
    });
  }

  return {
    available,
    tables,
    stats: { rows: total, kept, hiddenSkipped },
    keptRows,
    excludedCols,
    missing: { columns: missingColumns, conditions: missingConditions },
  };
}

function labelOf(range, context) {
  const named = (context.tables || []).find((t) => t.range.top === range.top && t.range.left === range.left
    && t.range.bottom === range.bottom && t.range.right === range.right);
  return named ? named.name : formatRange(range);
}

// Distinct values of a column within the ranges, for the value pickers:
// [{ value, count }] in natural order. Rows are those that pass the other
// conditions (`except` is the condition being edited), so the list shows
// what is still there to choose from.
function distinctValues(sheet, ranges, recipe, columnRef, except) {
  prepare(sheet);
  const date1904 = !!sheet.date1904;
  const parts = ranges.map((r) => readRange(sheet, r, recipe));
  const available = [];
  const seen = new Set();
  for (const p of parts) for (const c of p.columns) if (!seen.has(c.id)) { seen.add(c.id); available.push(c); }
  const col = findColumn(available, columnRef);
  if (!col) return [];
  const others = [];
  for (const cond of recipe.filter.conditions) {
    if (cond === except || (except && cond.id === except.id) || !cond.column) continue;
    const c = findColumn(available, cond.column);
    if (c && filter.isActive(cond, cond.values)) others.push({ cond, id: c.id });
  }
  const matchAll = recipe.filter.match !== 'any';
  const counts = new Map();
  for (const p of parts) {
    for (const row of p.rows) {
      if (others.length && matchAll && !others.every(({ cond, id }) => filter.test(cond, row.cells.get(id), cond.values, date1904))) continue;
      const v = row.cells.get(col.id);
      const text = v ? v.text : '';
      if (text === '') continue;
      counts.set(text, (counts.get(text) || 0) + 1);
    }
  }
  return [...counts.entries()]
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => a.value.localeCompare(b.value, undefined, { numeric: true, sensitivity: 'base' }));
}

module.exports = { extract, distinctValues, isoOfSerial };
