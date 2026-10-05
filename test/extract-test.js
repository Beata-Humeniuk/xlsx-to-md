'use strict';

const { check, eq, has, lacks } = require('./harness');
const { loadSpec } = require('./fixtures');
const { extract, distinctValues } = require('../src/core/extract');
const { renderMarkdown, escapeCell } = require('../src/core/markdown');
const R = require('../src/core/recipe');
const region = require('../src/core/region');

const wb = loadSpec({ hiddenRows: [4] });
const sheet = wb.loadSheet(0);
const TABLE = { top: 2, left: 0, bottom: 7, right: 8 };

function recipe(change) {
  const r = R.emptyRecipe();
  r.source.sheet = 'Messages';
  if (change) change(r);
  return R.normalize(r);
}

const cond = (name, op, values, ask = false) => ({ id: name + op, column: { name, letter: '' }, op, values, ask });

check('current region and data end', () => {
  eq(JSON.stringify(region.currentRegion(sheet, 3, 1)), JSON.stringify(TABLE));
  eq(region.dataEnd(sheet, 2, 0, 8), 7);
  eq(JSON.stringify(region.usedRange(sheet)), JSON.stringify({ top: 0, left: 0, bottom: 9, right: 8 }));
  eq(region.looksLikeHeader(sheet, TABLE), true);
  eq(region.looksLikeHeader(sheet, { top: 3, left: 0, bottom: 7, right: 8 }), false);
});

check('all columns, hidden row skipped', () => {
  const r = extract(sheet, [TABLE], recipe());
  eq(r.tables.length, 1);
  eq(r.tables[0].columns.map((c) => c.name).join(','), 'Field,Type,Required,Description,Status,Owner,System,Direction,Amount');
  eq(r.tables[0].rows.map((row) => row[0]).join(','), 'id,old_code,amount,items');
  eq(r.stats.hiddenSkipped, 1);
});

check('chosen columns in chosen order, by name', () => {
  const r = extract(sheet, [TABLE], recipe((x) => {
    x.columns = [
      { name: 'Description', letter: 'Z', include: true },
      { name: 'field', letter: 'Z', include: true },
      { name: 'Status', letter: 'E', include: false },
    ];
  }));
  eq(r.tables[0].columns.map((c) => c.name).join(','), 'Description,Field');
});

check('conditions, all and any', () => {
  const all = extract(sheet, [TABLE], recipe((x) => {
    x.filter.conditions = [cond('Status', 'neq', ['deleted']), cond('Required', 'eq', ['TRUE']), cond('System', 'eq', ['CRM'])];
  }));
  eq(all.tables[0].rows.map((row) => row[0]).join(','), 'id,items');
  eq(all.stats.kept, 2);
  eq([...all.keptRows].join(','), '3,7');
  const any = extract(sheet, [TABLE], recipe((x) => {
    x.filter.match = 'any';
    x.filter.conditions = [cond('System', 'eq', ['FUND']), cond('Field', 'eq', ['id'])];
  }));
  eq(any.tables[0].rows.map((row) => row[0]).join(','), 'id,amount');
});

check('a condition on a missing column is reported and skipped', () => {
  const r = extract(sheet, [TABLE], recipe((x) => {
    x.filter.conditions = [cond('Jira', 'eq', ['X'])];
  }));
  eq(r.stats.kept, 4);
  eq(r.missing.conditions.length, 1);
});

check('ask condition without a value keeps every row', () => {
  const r = extract(sheet, [TABLE], recipe((x) => {
    x.filter.conditions = [cond('System', 'eq', [], true)];
  }));
  eq(r.stats.kept, 4);
});

check('sorting', () => {
  const r = extract(sheet, [TABLE], recipe((x) => {
    x.sort = { column: { name: 'Amount', letter: 'I' }, dir: 'desc' };
  }));
  eq(r.tables[0].rows.map((row) => row[0]).join(','), 'amount,id,old_code,items');
});

check('hidden rows kept when asked', () => {
  const r = extract(sheet, [TABLE], recipe((x) => { x.options.skipHiddenRows = false; }));
  eq(r.stats.kept, 5);
});

check('no header row: columns by letter', () => {
  const r = extract(sheet, [{ top: 3, left: 0, bottom: 4, right: 1 }], recipe((x) => { x.source.headerRow = false; x.options.skipHiddenRows = false; }));
  eq(r.tables[0].columns.map((c) => c.id).join(','), 'c:A,c:B');
  const md = renderMarkdown(r, recipe((x) => { x.source.headerRow = false; }));
  has(md, '| A    | B      |');
});

check('merged cells fill their area', () => {
  const r = extract(sheet, [{ top: 0, left: 0, bottom: 0, right: 3 }], recipe((x) => { x.source.headerRow = false; }));
  eq(r.tables[0].rows[0].join('|'), 'Message specification|Message specification|Message specification|Message specification');
  const off = extract(sheet, [{ top: 0, left: 0, bottom: 0, right: 3 }], recipe((x) => { x.source.headerRow = false; x.options.fillMerged = false; }));
  eq(off.tables[0].rows[0].join('|'), 'Message specification|||');
});

check('several ranges: merged when the columns match, separate otherwise', () => {
  const top = { top: 2, left: 0, bottom: 3, right: 1 };
  const second = { top: 2, left: 0, bottom: 2, right: 1 };
  const same = extract(sheet, [top, second], recipe());
  eq(same.tables.length, 1);
  const other = extract(sheet, [top, { top: 2, left: 4, bottom: 3, right: 5 }], recipe());
  eq(other.tables.length, 2);
  const forced = extract(sheet, [top, { top: 2, left: 4, bottom: 3, right: 5 }], recipe((x) => { x.options.multiRange = 'merge'; }));
  eq(forced.tables.length, 1);
  eq(forced.tables[0].columns.length, 4);
});

check('values without number formats', () => {
  const r = extract(sheet, [{ top: 2, left: 8, bottom: 6, right: 8 }], recipe((x) => { x.options.values = 'raw'; }));
  eq(r.tables[0].rows.map((row) => row[0]).join(','), '10,7,1200');
});

check('distinct values follow the other conditions', () => {
  const r = recipe((x) => {
    x.filter.conditions = [cond('Status', 'neq', ['DELETED']), cond('Owner', 'eq', [], true)];
  });
  const values = distinctValues(sheet, [TABLE], r, { name: 'Owner' }, r.filter.conditions[1]);
  eq(values.map((v) => v.value + ':' + v.count).join(','), 'anna:1,ewa:1,piotr:1');
});

check('markdown table', () => {
  const r = recipe((x) => {
    x.columns = ['Field', 'Type', 'Description', 'Amount'].map((name) => ({ name, letter: '', include: true }));
  });
  const md = renderMarkdown(extract(sheet, [TABLE], r), r, { sheetName: 'Messages' });
  eq(md, [
    '| Field    | Type        | Description       | Amount |',
    '| -------- | ----------- | ----------------- | -----: |',
    '| id       | string      | Identifier        |     10 |',
    '| old_code | int         | Legacy            |      7 |',
    '| amount   | decimal     | Value \\| net      |   1200 |',
    '| items    | List\\<Item> | \\*Required\\* list |      3 |',
    '',
  ].join('\n'));
});

check('markdown title and sections', () => {
  const r = recipe((x) => {
    x.columns = ['Field', 'Type', 'Description'].map((name) => ({ name, letter: '', include: true }));
    x.options.layout = 'sections';
    x.options.title = 'filter';
    x.options.skipHiddenRows = false;
    x.filter.conditions = [cond('Field', 'eq', ['name'])];
  });
  const md = renderMarkdown(extract(sheet, [TABLE], r), r, { sheetName: 'Messages', filterName: 'Spec' });
  eq(md, '## Spec\n\n### name\n\n- **Type:** string\n- **Description:** Display name<br>shown in lists\n');
});

check('line breaks as spaces, no padding', () => {
  const r = recipe((x) => {
    x.columns = [{ name: 'Description', letter: '', include: true }];
    x.options.lineBreaks = 'space';
    x.options.padColumns = false;
    x.options.skipHiddenRows = false;
    x.filter.conditions = [cond('Field', 'eq', ['name'])];
  });
  const md = renderMarkdown(extract(sheet, [TABLE], r), r);
  eq(md, '| Description |\n| --- |\n| Display name shown in lists |\n');
});

check('escaping', () => {
  eq(escapeCell('a|b'), 'a\\|b');
  eq(escapeCell('List<String>'), 'List\\<String>');
  eq(escapeCell('a < b'), 'a < b');
  eq(escapeCell('field_name_x'), 'field_name_x');
  eq(escapeCell('_private'), '\\_private');
  eq(escapeCell('2*3'), '2\\*3');
  eq(escapeCell('C:\\temp'), 'C:\\temp');
  eq(escapeCell('one\ntwo'), 'one<br>two');
  lacks(escapeCell('ok'), '\\');
});

require('./harness').done('extract');
