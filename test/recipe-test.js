'use strict';

const { check, checkAsync, eq } = require('./harness');
const R = require('../src/core/recipe');
const { FilterStore, summary, conditionsText } = require('../src/filters');
const { forLanguage } = require('../src/nls');

function base() {
  const r = R.emptyRecipe();
  r.source = { sheet: 'S', headerRow: true, ranges: [{ ref: 'A1:C9', table: null, headers: ['A', 'B', 'C'], toEnd: true }] };
  r.columns = [{ name: 'A', letter: 'A', include: true }, { name: 'B', letter: 'B', include: false }];
  r.filter.conditions = [
    { id: 'x', column: { name: 'A', letter: 'A' }, op: 'neq', values: ['DELETED'], ask: false },
    { id: 'y', column: { name: 'B', letter: 'B' }, op: 'eq', values: ['CRM'], ask: true },
  ];
  return R.normalize(r);
}

check('normalize repairs partial data', () => {
  const r = R.normalize({ columns: [{ name: 'X' }, null, {}], filter: { conditions: [{ op: 'eq' }] }, options: { layout: 'sections', bogus: 1 } });
  eq(r.columns.length, 1);
  eq(r.columns[0].include, true);
  eq(r.filter.conditions[0].id, 'c1');
  eq(r.options.layout, 'sections');
  eq('bogus' in r.options, false);
  eq(r.options.skipEmptyRows, true);
});

check('saving leaves out the values given when used', () => {
  const s = R.forSaving(base());
  eq(s.filter.conditions[0].values.join(), 'DELETED');
  eq(s.filter.conditions[1].values.length, 0);
});

check('modified: ask values do not count, fixed values do', () => {
  const saved = R.forSaving(base());
  const cur = base();
  eq(R.isModified(cur, saved, false), false);
  cur.filter.conditions[1].values = ['FUND'];
  eq(R.isModified(cur, saved, false), false);
  cur.filter.conditions[0].values = ['OLD'];
  eq(R.isModified(cur, saved, false), true);
});

check('modified: column order and choice count, excluded extras do not', () => {
  const saved = R.forSaving(base());
  const cur = base();
  cur.columns.push({ name: 'C', letter: 'C', include: false });
  eq(R.isModified(cur, saved, false), false);
  cur.columns[2].include = true;
  eq(R.isModified(cur, saved, false), true);
  const swapped = base();
  swapped.columns[1].include = true;
  const saved2 = R.forSaving(swapped);
  swapped.columns.reverse();
  eq(R.isModified(swapped, saved2, false), true);
});

check('modified: options and the ask switch', () => {
  const saved = R.forSaving(base());
  const cur = base();
  cur.options.layout = 'sections';
  eq(R.isModified(cur, saved, false), true);
  const ask = base();
  ask.filter.conditions[0].ask = true;
  eq(R.isModified(ask, saved, false), true);
});

check('ask conditions', () => {
  eq(R.askConditions(base()).map((c) => c.id).join(), 'y');
});

check('reconciling columns keeps choices and adds new ones', () => {
  const prev = [{ name: 'B', letter: 'B', include: false }, { name: 'A', letter: 'A', include: true }];
  const avail = [{ id: 'h:a', name: 'A', letter: 'C' }, { id: 'h:b', name: 'B', letter: 'D' }, { id: 'h:new', name: 'New', letter: 'E' }];
  const user = R.reconcileColumns(prev, avail, true);
  eq(user.columns.map((c) => c.name + ':' + c.include).join(','), 'B:false,A:true,New:true');
  const filter = R.reconcileColumns(prev.concat([{ name: 'Gone', letter: 'Z', include: true }]), avail, false);
  eq(filter.columns.map((c) => c.name + ':' + c.include).join(','), 'B:false,A:true,New:false');
  eq(filter.added.map((c) => c.name).join(), 'New');
  eq(filter.missing.map((c) => c.name).join(), 'Gone');
});

function memoryState() {
  const data = new Map();
  return {
    get: (k, d) => (data.has(k) ? data.get(k) : d),
    update: async (k, v) => { data.set(k, v); },
    setKeysForSync: () => {},
  };
}

checkAsync('filter store: save, rename, remove, remembered values', async () => {
  const store = new FilterStore(memoryState());
  let changes = 0;
  store.onDidChange(() => changes++);
  const f = await store.save({ name: ' Spec ', recipe: base() });
  eq(store.list().length, 1);
  eq(store.get(f.id).name, 'Spec');
  eq(store.get(f.id).recipe.filter.conditions[1].values.length, 0);
  eq(store.findByName('spec').id, f.id);
  await store.save({ id: f.id, name: 'Spec', recipe: base() });
  eq(store.list().length, 1);
  await store.rename(f.id, 'Other');
  eq(store.get(f.id).name, 'Other');
  await store.rememberValues(f.id, { y: ['CRM'] });
  eq(store.lastValues(f.id).y[0], 'CRM');
  await store.remove(f.id);
  eq(store.list().length, 0);
  eq(Object.keys(store.lastValues(f.id)).length, 0);
  eq(changes, 4);
});

check('filter in plain words', () => {
  const nls = forLanguage('pl');
  const f = { name: 'Spec', recipe: base() };
  eq(summary(f, nls), 'arkusz S · 1 kolumna · 2 warunki, 1 podawane przy użyciu');
  eq(conditionsText(f, nls), 'A ≠ DELETED · B = ?');
});

require('./harness').done('recipe');
