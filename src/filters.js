'use strict';

// Saved filters. They are global: kept in the extension's global state (and
// synced with Settings Sync), never per workspace, so the same filter works
// in any project and on any workbook.

const { normalize, forSaving } = require('./core/recipe');
const { refLabel } = require('./core/columns');

const KEY = 'xlsxToMd.savedFilters';
const VALUES_KEY = 'xlsxToMd.lastValues';

function newId() {
  return 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

class FilterStore {
  constructor(globalState) {
    this.state = globalState;
    this.listeners = new Set();
    if (typeof globalState.setKeysForSync === 'function') globalState.setKeysForSync([KEY]);
  }

  list() {
    const raw = this.state.get(KEY, []);
    return (Array.isArray(raw) ? raw : [])
      .filter((f) => f && f.id && f.name)
      .map((f) => ({ id: String(f.id), name: String(f.name), updated: f.updated || 0, recipe: normalize(f.recipe) }))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  }

  get(id) {
    return this.list().find((f) => f.id === id) || null;
  }

  findByName(name) {
    const key = String(name).trim().toLowerCase();
    return this.list().find((f) => f.name.trim().toLowerCase() === key) || null;
  }

  async write(filters) {
    await this.state.update(KEY, filters.map((f) => ({ id: f.id, name: f.name, updated: f.updated, recipe: f.recipe })));
    for (const l of this.listeners) l();
  }

  // Creates a filter, or replaces the one with `id`. Returns the stored filter.
  async save({ id, name, recipe }) {
    const filters = this.list();
    const entry = { id: id || newId(), name: String(name).trim(), updated: Date.now(), recipe: forSaving(recipe) };
    const at = filters.findIndex((f) => f.id === entry.id);
    if (at >= 0) filters[at] = entry;
    else filters.push(entry);
    await this.write(filters);
    return entry;
  }

  async rename(id, name) {
    const filters = this.list();
    const f = filters.find((x) => x.id === id);
    if (!f) return;
    f.name = String(name).trim();
    f.updated = Date.now();
    await this.write(filters);
  }

  async remove(id) {
    await this.write(this.list().filter((f) => f.id !== id));
    const values = { ...this.state.get(VALUES_KEY, {}) };
    if (values[id]) {
      delete values[id];
      await this.state.update(VALUES_KEY, values);
    }
  }

  // The values last given for a filter's "ask when used" conditions.
  lastValues(id) {
    const all = this.state.get(VALUES_KEY, {});
    return (all && all[id]) || {};
  }

  async rememberValues(id, values) {
    const all = { ...this.state.get(VALUES_KEY, {}) };
    all[id] = values;
    await this.state.update(VALUES_KEY, all);
  }

  onDidChange(listener) {
    this.listeners.add(listener);
    return { dispose: () => this.listeners.delete(listener) };
  }
}

// A filter in plain words, for lists: "sheet Messages · 4 columns · 3 conditions, 2 given when used".
function summary(filter, nls) {
  const r = filter.recipe;
  const parts = [];
  if (r.source.sheet) parts.push(nls.t('summary.sheet', { sheet: r.source.sheet }));
  const included = r.columns.filter((c) => c.include).length;
  if (included) parts.push(nls.t('summary.columns', { count: included, columns: nls.plural('plural.column', included) }));
  const conds = r.filter.conditions.filter((c) => c.column);
  if (!conds.length) parts.push(nls.t('summary.noConditions'));
  else {
    let text = nls.t('summary.conditions', { count: conds.length, conditions: nls.plural('plural.condition', conds.length) });
    const ask = conds.filter((c) => c.ask).length;
    if (ask) text += ', ' + nls.t('summary.ask', { count: ask });
    parts.push(text);
  }
  return parts.join(' · ');
}

// The conditions in plain words: "Status ≠ DELETED · System = ?".
function conditionsText(filter, nls) {
  const r = filter.recipe;
  const sep = r.filter.match === 'any' ? ' | ' : ' · ';
  return r.filter.conditions.filter((c) => c.column).map((c) => {
    const op = nls.t('op.' + c.op);
    if (c.op === 'empty' || c.op === 'notEmpty') return refLabel(c.column, nls) + ' ' + op;
    const value = c.ask ? nls.t('summary.askValue') : c.values.join(', ');
    return refLabel(c.column, nls) + ' ' + op + ' ' + value;
  }).join(sep);
}

module.exports = { FilterStore, summary, conditionsText };
