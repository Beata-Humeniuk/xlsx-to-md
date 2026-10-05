'use strict';

// Row conditions. A condition is { id, column: { name, letter }, op,
// values: [], ask }. Several values mean "any of them" (for ≠ and "does not
// contain": "none of them"). A condition with no value yet — typically one
// whose value is given when the filter is used and was left blank — does not
// filter anything.

const OPERATORS = [
  { op: 'eq', types: ['text', 'number', 'date', 'bool'], multi: true },
  { op: 'neq', types: ['text', 'number', 'date', 'bool'], multi: true },
  { op: 'contains', types: ['text', 'number', 'date'], multi: true },
  { op: 'notContains', types: ['text', 'number', 'date'], multi: true },
  { op: 'startsWith', types: ['text', 'number'], multi: true },
  { op: 'endsWith', types: ['text', 'number'], multi: true },
  { op: 'gt', types: ['number', 'date', 'text'], multi: false },
  { op: 'gte', types: ['number', 'date', 'text'], multi: false },
  { op: 'lt', types: ['number', 'date', 'text'], multi: false },
  { op: 'lte', types: ['number', 'date', 'text'], multi: false },
  { op: 'empty', types: ['text', 'number', 'date', 'bool'], multi: false, noValue: true },
  { op: 'notEmpty', types: ['text', 'number', 'date', 'bool'], multi: false, noValue: true },
];

const BY_OP = new Map(OPERATORS.map((o) => [o.op, o]));

function operator(op) {
  return BY_OP.get(op) || BY_OP.get('eq');
}

// Operators that make sense for a column of the given type, most useful first.
// Comparisons (> < ≥ ≤) are offered for numbers and dates, not for text.
function operatorsFor(type) {
  return OPERATORS.filter((o) => {
    if (type === 'text' && ['gt', 'gte', 'lt', 'lte'].includes(o.op)) return false;
    return o.types.includes(type || 'text');
  }).map((o) => o.op);
}

const TRUE_WORDS = new Set(['true', 'prawda', 'tak', 'yes', 'y', 't']);
const FALSE_WORDS = new Set(['false', 'fałsz', 'falsz', 'nie', 'no', 'n', 'f']);

function fold(s) {
  return String(s == null ? '' : s).normalize('NFKC').trim().toLowerCase();
}

function boolOf(s) {
  const f = fold(s);
  if (TRUE_WORDS.has(f)) return true;
  if (FALSE_WORDS.has(f)) return false;
  return null;
}

// "1 234,5", "1,234.50", "-3", "12%" → number; null when it is not a number.
function parseNumber(text) {
  let s = String(text == null ? '' : text).trim().replace(/[\s\u00a0\u202f]/g, '');
  if (!s) return null;
  let scale = 1;
  if (s.endsWith('%')) {
    scale = 0.01;
    s = s.slice(0, -1);
  }
  if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
  else if (/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^[-+]?\d*,\d+$/.test(s)) s = s.replace(',', '.');
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
  const n = Number(s) * scale;
  return Number.isFinite(n) ? n : null;
}

// "2026-03-15", "2026-03-15 10:30", "15.03.2026", "15/03/2026" → Excel serial
// in the given date system; null when it is not a date.
function parseDate(text, date1904) {
  const s = String(text == null ? '' : text).trim();
  let y;
  let mo;
  let d;
  let rest = '';
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(.*)$/.exec(s);
  if (m) {
    [y, mo, d, rest] = [+m[1], +m[2], +m[3], m[4]];
  } else if ((m = /^(\d{1,2})[./](\d{1,2})[./](\d{4})(.*)$/.exec(s))) {
    [d, mo, y, rest] = [+m[1], +m[2], +m[3], m[4]];
  } else {
    return null;
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  let seconds = 0;
  const t = /^[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$/.exec(rest);
  if (t) seconds = +t[1] * 3600 + +t[2] * 60 + (t[3] ? +t[3] : 0);
  else if (rest.trim()) return null;
  const ms = Date.UTC(y, mo - 1, d);
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  let serial = (ms - epoch) / 86400000;
  if (!date1904 && serial < 61) serial -= 1;
  return serial + seconds / 86400;
}

// The column type the values suggest: 'number', 'date', 'bool' or 'text'.
function columnType(cells) {
  let n = 0;
  let dates = 0;
  let nums = 0;
  let bools = 0;
  for (const c of cells) {
    if (!c || c.text === '') continue;
    n++;
    if (c.kind === 'd') dates++;
    else if (c.kind === 'n') nums++;
    else if (c.kind === 'b' || boolOf(c.text) !== null) bools++;
    else if (parseNumber(c.text) !== null) nums++;
  }
  if (!n) return 'text';
  if (dates / n >= 0.8) return 'date';
  if (nums / n >= 0.8) return 'number';
  if (bools / n >= 0.8) return 'bool';
  return 'text';
}

function cellNumber(cell) {
  if (cell.num != null && cell.kind !== 'b') return cell.num;
  return parseNumber(cell.text);
}

// Compares a cell to a value given as text: numerically, as dates or in
// natural text order. Returns <0, 0, >0, or null when they cannot be ordered.
function compare(cell, value, date1904) {
  if (cell.kind === 'd' && cell.num != null) {
    const v = parseDate(value, date1904);
    if (v !== null) {
      // A date without time matches the whole day.
      const day = /\d:\d/.test(value) ? cell.num : Math.floor(cell.num);
      return day - v;
    }
  }
  const a = cellNumber(cell);
  const b = parseNumber(value);
  if (a !== null && b !== null) return a - b;
  if (cell.text === '') return null;
  return fold(cell.text).localeCompare(fold(value), undefined, { numeric: true });
}

function equals(cell, value, date1904) {
  const text = fold(cell.text);
  const v = fold(value);
  if (text === v) return true;
  if (v === '' || text === '') return false;
  const cb = cell.kind === 'b' ? cell.num === 1 : boolOf(cell.text);
  const vb = boolOf(value);
  if (cb !== null && vb !== null) return cb === vb;
  if (cell.kind === 'd' && cell.num != null) {
    const d = parseDate(value, date1904);
    if (d !== null) return /\d:\d/.test(value) ? Math.abs(cell.num - d) < 1 / 86400 : Math.floor(cell.num) === Math.floor(d);
  }
  const a = cellNumber(cell);
  const b = parseNumber(value);
  return a !== null && b !== null && Math.abs(a - b) < 1e-9 * Math.max(1, Math.abs(a));
}

function hasValue(values) {
  return Array.isArray(values) && values.some((v) => String(v).trim() !== '');
}

function activeValues(values) {
  return (values || []).map(String).filter((v) => v.trim() !== '');
}

// Whether a condition has what it needs to filter.
function isActive(cond, values) {
  const op = operator(cond.op);
  return op.noValue || hasValue(values);
}

function test(cond, cell, values, date1904) {
  const c = cell || { text: '', kind: '', num: null };
  const vals = activeValues(values);
  const text = fold(c.text);
  switch (cond.op) {
    case 'empty': return text === '';
    case 'notEmpty': return text !== '';
    case 'eq': return vals.some((v) => equals(c, v, date1904));
    case 'neq': return !vals.some((v) => equals(c, v, date1904));
    case 'contains': return vals.some((v) => text.includes(fold(v)));
    case 'notContains': return !vals.some((v) => text.includes(fold(v)));
    case 'startsWith': return vals.some((v) => text.startsWith(fold(v)));
    case 'endsWith': return vals.some((v) => text.endsWith(fold(v)));
    case 'gt': case 'gte': case 'lt': case 'lte': {
      const r = compare(c, vals[0], date1904);
      if (r === null) return false;
      if (cond.op === 'gt') return r > 0;
      if (cond.op === 'gte') return r >= 0;
      if (cond.op === 'lt') return r < 0;
      return r <= 0;
    }
    default: return true;
  }
}

module.exports = {
  OPERATORS, operator, operatorsFor, columnType, parseNumber, parseDate, boolOf,
  isActive, test, activeValues, fold,
};
