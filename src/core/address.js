'use strict';

// A1-style cell addresses. Rows and columns are 0-based inside the extension;
// only the text the user sees ("B4", "A4:G120") is 1-based like in Excel.

function columnLetter(index) {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function columnIndex(letters) {
  let n = 0;
  for (const ch of String(letters).toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COLUMN = /^\$?([A-Za-z]{1,3})$/;
const ROW = /^\$?(\d{1,7})$/;

// "B4" → { row: 3, col: 1 }; null when it is not a cell address.
function parseCell(text) {
  const m = CELL.exec(String(text).trim());
  if (!m) return null;
  const row = parseInt(m[2], 10) - 1;
  const col = columnIndex(m[1]);
  if (row < 0 || col < 0) return null;
  return { row, col };
}

function formatCell(row, col) {
  return columnLetter(col) + (row + 1);
}

// "A4:G120", "B7", "C:E" (whole columns) or "3:9" (whole rows) → a range
// { top, left, bottom, right } with inclusive ends. Whole rows and columns are
// clipped to `limits` ({ rows, cols }), the size of the sheet.
function parseRange(text, limits) {
  const parts = String(text).trim().split(':');
  if (parts.length === 1) {
    const c = parseCell(parts[0]);
    return c ? { top: c.row, left: c.col, bottom: c.row, right: c.col } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1]);
  if (a && b) return normalizeRange({ top: a.row, left: a.col, bottom: b.row, right: b.col });
  const maxRow = Math.max(0, ((limits && limits.rows) || 1) - 1);
  const maxCol = Math.max(0, ((limits && limits.cols) || 1) - 1);
  const ca = COLUMN.exec(parts[0].trim());
  const cb = COLUMN.exec(parts[1].trim());
  if (ca && cb) {
    return normalizeRange({ top: 0, left: columnIndex(ca[1]), bottom: maxRow, right: columnIndex(cb[1]) });
  }
  const ra = ROW.exec(parts[0].trim());
  const rb = ROW.exec(parts[1].trim());
  if (ra && rb) {
    return normalizeRange({ top: parseInt(ra[1], 10) - 1, left: 0, bottom: parseInt(rb[1], 10) - 1, right: maxCol });
  }
  return null;
}

// Several ranges separated by commas or semicolons, as typed into the name box.
function parseRangeList(text, limits) {
  const out = [];
  for (const piece of String(text).split(/[,;]/)) {
    if (!piece.trim()) continue;
    const r = parseRange(piece, limits);
    if (!r || r.top < 0 || r.left < 0) return null;
    out.push(r);
  }
  return out;
}

function normalizeRange(r) {
  return {
    top: Math.min(r.top, r.bottom),
    left: Math.min(r.left, r.right),
    bottom: Math.max(r.top, r.bottom),
    right: Math.max(r.left, r.right),
  };
}

function formatRange(r) {
  const a = formatCell(r.top, r.left);
  if (r.top === r.bottom && r.left === r.right) return a;
  return a + ':' + formatCell(r.bottom, r.right);
}

function contains(r, row, col) {
  return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

function sameRange(a, b) {
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

module.exports = {
  columnLetter, columnIndex, parseCell, formatCell, parseRange, parseRangeList,
  normalizeRange, formatRange, contains, sameRange,
};
