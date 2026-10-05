'use strict';

const { check, eq } = require('./harness');
const { xlsx } = require('./xlsx-builder');
const { openWorkbook } = require('../src/workbook');
const { loadSpec, specRows, HEAD } = require('./fixtures');
const { resolveSource, describeRange } = require('../src/core/resolve');
const R = require('../src/core/recipe');

const TABLE = { top: 2, left: 0, bottom: 7, right: 8 };

function savedFrom(wb, range, change) {
  const sheet = wb.loadSheet(0);
  const r = R.emptyRecipe();
  r.source.sheet = sheet.name;
  r.source.ranges = [describeRange(sheet, range, true)];
  if (change) change(r);
  return R.forSaving(r);
}

const same = (a, b) => eq(JSON.stringify(a), JSON.stringify(b));

check('describing a range', () => {
  const d = describeRange(loadSpec().loadSheet(0), TABLE, true);
  eq(d.ref, 'A3:I8');
  eq(d.table, null);
  eq(d.headers.join(','), HEAD.join(','));
  eq(d.toEnd, true);
  eq(describeRange(loadSpec().loadSheet(0), { top: 2, left: 0, bottom: 5, right: 8 }, true).toEnd, false);
});

check('same file: same place', () => {
  const wb = loadSpec();
  const hit = resolveSource(wb, savedFrom(wb, TABLE));
  eq(hit.sheetIndex, 0);
  same(hit.ranges, [TABLE]);
  eq(hit.notes.length, 0);
});

check('rows added and columns moved: found by headers', () => {
  const wb = loadSpec();
  const saved = savedFrom(wb, TABLE);
  const rows = specRows().map((r) => r.slice());
  // Two rows of preamble more, a column inserted in front, three more data rows.
  const moved = [['Version 4'], []].concat(rows.slice(0, 8).map((r, i) => (i >= 2 ? ['x'].concat(r) : r)))
    .concat([['x', 'extra1', 'string', true, '', 'OK', 'a', 'CRM', 'IN', 1], ['x', 'extra2', 'string', true, '', 'OK', 'a', 'CRM', 'IN', 1]]);
  moved[4][0] = 'Jira';
  const wb2 = openWorkbook(xlsx([{ name: 'Messages', rows: moved }]));
  const hit = resolveSource(wb2, saved);
  same(hit.ranges, [{ top: 4, left: 0, bottom: 11, right: 9 }]);
});

check('fixed number of rows when the range did not reach the end', () => {
  const wb = loadSpec();
  const saved = savedFrom(wb, { top: 2, left: 0, bottom: 4, right: 8 });
  const hit = resolveSource(wb, saved);
  same(hit.ranges, [{ top: 2, left: 0, bottom: 4, right: 8 }]);
});

check('sheet renamed: found by headers', () => {
  const wb = loadSpec();
  const saved = savedFrom(wb, TABLE);
  const wb2 = openWorkbook(xlsx([{ name: 'Intro', rows: [['hello']] }, { name: 'Komunikaty', rows: specRows() }]));
  const hit = resolveSource(wb2, saved);
  eq(hit.sheetIndex, 1);
  eq(hit.notes[0].code, 'sheetByHeaders');
});

check('Excel table found by name', () => {
  const wb = openWorkbook(xlsx([{ name: 'S', rows: [[], ['A', 'B'], [1, 2], [3, 4]], tables: [{ name: 'Codes', ref: 'A2:B4' }] }]));
  const saved = savedFrom(wb, { top: 1, left: 0, bottom: 3, right: 1 });
  eq(saved.source.ranges[0].table, 'Codes');
  const wb2 = openWorkbook(xlsx([{ name: 'S', rows: [['A', 'B'], [1, 2], [3, 4], [5, 6]], tables: [{ name: 'Codes', ref: 'A1:B4' }] }]));
  same(resolveSource(wb2, saved).ranges, [{ top: 0, left: 0, bottom: 3, right: 1 }]);
});

check('headers not found: saved address, with a note', () => {
  const wb = loadSpec();
  const saved = savedFrom(wb, TABLE);
  const wb2 = openWorkbook(xlsx([{ name: 'Messages', rows: [['a', 'b'], ['c', 'd']] }]));
  const hit = resolveSource(wb2, saved);
  eq(hit.notes.some((n) => n.code === 'headersNotFound'), true);
  eq(hit.ranges.length, 1);
});

check('missing sheet without headers: first visible sheet', () => {
  const r = R.emptyRecipe();
  r.source = { sheet: 'Gone', headerRow: false, ranges: [{ ref: 'A1:B2', table: null, headers: null, toEnd: false }] };
  const wb = openWorkbook(xlsx([{ name: 'Hidden', hidden: true, rows: [['x']] }, { name: 'Shown', rows: [['y']] }]));
  const hit = resolveSource(wb, R.normalize(r));
  eq(hit.sheetIndex, 1);
  eq(hit.notes[0].code, 'sheetMissing');
});

require('./harness').done('resolve');
