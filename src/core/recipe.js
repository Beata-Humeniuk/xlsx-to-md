'use strict';

// A recipe is everything the user clicked together to get their Markdown:
// which sheet and ranges, which columns in what order, which row conditions,
// sorting and output options. A saved filter is a named recipe. Recipes are
// stored internally as plain objects; the user never sees or edits them.
//
// {
//   source:  { sheet, headerRow, ranges: [{ ref, table, headers, toEnd }] },
//   columns: [{ name, letter, include }],           // output order
//   filter:  { match: 'all' | 'any', conditions: [{ id, column, op, values, ask }] },
//   sort:    { column, dir: 'asc' | 'desc' },
//   options: { … see DEFAULT_OPTIONS }
// }
//
// A condition with `ask: true` has its value given each time the filter is
// used; its `values` are not part of the saved filter.

const { refId } = require('./columns');

const DEFAULT_OPTIONS = {
  skipHiddenRows: true,
  skipHiddenColumns: true,
  skipEmptyRows: true,
  dropEmptyColumns: false,
  fillMerged: true,
  values: 'formatted', // 'formatted' (as Excel shows them) | 'raw'
  layout: 'table', // 'table' | 'sections'
  title: 'none', // 'none' | 'sheet' | 'filter' | 'custom'
  titleText: '',
  multiRange: 'auto', // 'auto' | 'merge' | 'separate'
  lineBreaks: 'br', // 'br' | 'space'
  alignNumbers: true,
  padColumns: true,
};

function emptyRecipe() {
  return {
    source: { sheet: null, headerRow: true, ranges: [] },
    columns: [],
    filter: { match: 'all', conditions: [] },
    sort: { column: null, dir: 'asc' },
    options: { ...DEFAULT_OPTIONS },
  };
}

function str(v) {
  return v == null ? null : String(v);
}

function normRef(ref) {
  if (!ref || typeof ref !== 'object') return null;
  const name = ref.name == null || String(ref.name).trim() === '' ? null : String(ref.name);
  const letter = ref.letter == null ? '' : String(ref.letter).toUpperCase();
  if (!name && !letter) return null;
  return { name, letter };
}

// A recipe in its complete, well-formed shape (old or partial data repaired).
function normalize(r) {
  const base = emptyRecipe();
  if (!r || typeof r !== 'object') return base;
  const src = r.source || {};
  base.source = {
    sheet: str(src.sheet),
    headerRow: src.headerRow !== false,
    ranges: (Array.isArray(src.ranges) ? src.ranges : []).filter((x) => x && x.ref).map((x) => ({
      ref: String(x.ref),
      table: x.table ? String(x.table) : null,
      headers: Array.isArray(x.headers) ? x.headers.map((h) => (h == null ? null : String(h))) : null,
      toEnd: !!x.toEnd,
    })),
  };
  base.columns = (Array.isArray(r.columns) ? r.columns : []).map((c) => {
    const ref = normRef(c);
    return ref ? { ...ref, include: c.include !== false } : null;
  }).filter(Boolean);
  const f = r.filter || {};
  base.filter = {
    match: f.match === 'any' ? 'any' : 'all',
    conditions: (Array.isArray(f.conditions) ? f.conditions : []).map((c, i) => ({
      id: c && c.id ? String(c.id) : 'c' + (i + 1),
      column: normRef(c && c.column),
      op: c && c.op ? String(c.op) : 'eq',
      values: Array.isArray(c && c.values) ? c.values.map(String) : [],
      ask: !!(c && c.ask),
    })),
  };
  const s = r.sort || {};
  base.sort = { column: normRef(s.column), dir: s.dir === 'desc' ? 'desc' : 'asc' };
  base.options = { ...DEFAULT_OPTIONS };
  const o = r.options || {};
  for (const k of Object.keys(DEFAULT_OPTIONS)) {
    if (o[k] !== undefined && typeof o[k] === typeof DEFAULT_OPTIONS[k]) base.options[k] = o[k];
  }
  return base;
}

// What a saved filter stores: the recipe without the values that are given
// when the filter is used.
function forSaving(recipe) {
  const r = normalize(recipe);
  for (const c of r.filter.conditions) if (c.ask) c.values = [];
  return r;
}

// The parts of a recipe that decide whether it still matches a saved filter.
// The source is compared only when the user changed the selection — applying
// a filter to another file finds the data at other addresses, and that is
// not a change to the filter.
function signature(recipe, withSource) {
  const r = forSaving(recipe);
  // Only the chosen columns and their order count; columns left out are
  // remembered for convenience but do not change what the filter produces.
  const cols = r.columns.filter((c) => c.include).map((c) => refId(c));
  const conds = r.filter.conditions.map((c) => [refId(c.column), c.op, c.ask, c.ask ? [] : c.values]);
  const parts = {
    headerRow: r.source.headerRow,
    cols,
    match: r.filter.match,
    conds,
    sort: [refId(r.sort.column), r.sort.dir],
    options: r.options,
  };
  if (withSource) {
    parts.source = {
      sheet: r.source.sheet,
      ranges: r.source.ranges.map((x) => [x.table || x.ref, x.toEnd]),
    };
  }
  return JSON.stringify(parts);
}

function isModified(current, saved, sourceChanged) {
  return signature(current, sourceChanged) !== signature(saved, sourceChanged);
}

function askConditions(recipe) {
  return normalize(recipe).filter.conditions.filter((c) => c.ask && c.column && !['empty', 'notEmpty'].includes(c.op));
}

// Keeps the user's column choices and order for the columns still there and
// adds the new ones at the end — included when the user is choosing columns
// by hand, left out when they come with a saved filter (it did not ask for them).
function reconcileColumns(previous, available, includeNew) {
  const out = [];
  const have = new Set();
  const avail = new Map(available.map((c) => [c.id, c]));
  for (const p of previous || []) {
    const id = refId(p);
    const a = avail.get(id) || (p.name ? null : available.find((c) => !c.name && c.letter === p.letter));
    if (!a || have.has(a.id)) continue;
    have.add(a.id);
    out.push({ name: a.name, letter: a.letter, include: p.include !== false });
  }
  const missing = (previous || []).filter((p) => {
    const id = refId(p);
    return !avail.has(id) && !(p.name == null && available.some((c) => !c.name && c.letter === p.letter));
  });
  const added = [];
  for (const a of available) {
    if (have.has(a.id)) continue;
    have.add(a.id);
    out.push({ name: a.name, letter: a.letter, include: !!includeNew });
    added.push(a);
  }
  return { columns: out, missing, added };
}

let counter = 0;
function newConditionId() {
  counter = (counter + 1) % 1e6;
  return 'c' + Date.now().toString(36) + counter.toString(36);
}

module.exports = {
  DEFAULT_OPTIONS, emptyRecipe, normalize, forSaving, signature, isModified,
  askConditions, reconcileColumns, newConditionId,
};
