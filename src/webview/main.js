'use strict';

// The workbook editor's page: toolbar with saved filters and export, the
// sheet grid, the side panel (data, columns, row conditions, options) and the
// Markdown preview. All choices live here; the extension reads the file,
// keeps the saved filters and writes the export.

/* global acquireVsCodeApi */
const vscode = acquireVsCodeApi();
const { fromMessages } = require('../nls');
const { formatRange, parseRangeList } = require('../core/address');
const R = require('../core/recipe');
const { extract, distinctValues } = require('../core/extract');
const { renderMarkdown } = require('../core/markdown');
const F = require('../core/filter');
const region = require('../core/region');
const { describeRange, headerRowGuess } = require('../core/resolve');
const { refId, refLabel, findColumn } = require('../core/columns');
const { Grid } = require('./grid');
const { valuePicker, closePopup } = require('./picker');

const PREVIEW_ROWS = 300;

let T = null; // { t, plural }
const S = {
  fileName: '',
  sheets: [],
  sheetIndex: 0,
  data: new Map(),
  waiting: new Map(),
  ranges: [],
  activeRange: -1,
  recipe: R.emptyRecipe(),
  headerManual: false,
  filters: [],
  active: null, // { id, name, baseline, baseRanges, baseSheet, sourceChanged }
  notes: [],
  columnNotes: { added: [], missing: [] },
  result: null,
  markdown: '',
  ui: Object.assign({
    preview: 'code',
    sections: { source: true, columns: true, conditions: true, options: false },
    sideWidth: 380,
    previewHeight: 260,
  }, (vscode.getState() || {}).ui || {}),
};

const $ = {};
let grid = null;

function saveUi() {
  vscode.setState({ ui: S.ui });
}

function post(message) {
  vscode.postMessage(message);
}

function clone(x) {
  return JSON.parse(JSON.stringify(x));
}

// --- DOM helper -----------------------------------------------------------

function h(tag, props, ...children) {
  const e = document.createElement(tag);
  if (props) {
    for (const [k, v] of Object.entries(props)) {
      if (v === undefined || v === null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'value') e.value = v;
      else if (k === 'checked') e.checked = !!v;
      else if (k === 'style') e.setAttribute('style', v);
      else e.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  }
  return e;
}

// replaceChildren without the null/false placeholders of optional parts.
function fill(parent, ...children) {
  parent.replaceChildren(...children.flat().filter((c) => c !== null && c !== undefined && c !== false));
}

function select(options, value, onChange, cls) {
  const s = h('select', { class: cls || '', onChange: () => onChange(s.value) });
  for (const o of options) {
    if (o.group) {
      const g = h('optgroup', { label: o.group });
      for (const x of o.items) g.appendChild(h('option', { value: x.value, text: x.label }));
      s.appendChild(g);
    } else {
      s.appendChild(h('option', { value: o.value, text: o.label, disabled: o.disabled }));
    }
  }
  s.value = value;
  return s;
}

function checkbox(label, checked, onChange, title) {
  const input = h('input', { type: 'checkbox', checked, onChange: () => onChange(input.checked) });
  return h('label', { class: 'check-row', title }, input, h('span', { text: label }));
}

// --- sheets ---------------------------------------------------------------

function unpackSheet(p) {
  const kind = p.kinds.map((row) => Array.from(row, (ch) => (ch === '.' ? '' : ch)));
  const num = p.text.map((row) => new Array(row.length).fill(null));
  for (let i = 0; i < p.nums.length; i += 3) num[p.nums[i]][p.nums[i + 1]] = p.nums[i + 2];
  return {
    name: p.name,
    hidden: p.hidden,
    rows: p.rows,
    cols: p.cols,
    text: p.text,
    kind,
    kinds: p.kinds,
    num,
    hiddenRows: p.hiddenRows,
    hiddenCols: p.hiddenCols,
    merges: p.merges,
    tables: p.tables,
    date1904: p.date1904,
  };
}

function loadSheet(index) {
  if (S.data.has(index)) return Promise.resolve(S.data.get(index));
  return new Promise((resolve) => {
    if (!S.waiting.has(index)) {
      S.waiting.set(index, []);
      post({ type: 'loadSheet', index });
    }
    S.waiting.get(index).push(resolve);
  });
}

function sheet() {
  return S.data.get(S.sheetIndex) || null;
}

async function switchSheet(index) {
  if (index === S.sheetIndex && sheet()) return;
  await loadSheet(index);
  S.sheetIndex = index;
  S.ranges = [];
  S.activeRange = -1;
  S.headerManual = false;
  grid.setSheet(sheet());
  changed({ source: true, user: true });
}

// --- computing the output ------------------------------------------------

function context() {
  const sh = sheet();
  return { sheetName: sh ? sh.name : '', filterName: S.active ? S.active.name : '', tables: sh ? sh.tables : [] };
}

function syncSource(includeNew) {
  const sh = sheet();
  S.recipe.source.sheet = sh ? sh.name : null;
  if (!sh || !S.ranges.length) return;
  if (!S.headerManual) S.recipe.source.headerRow = headerRowGuess(sh, S.ranges);
  const probe = extract(sh, S.ranges, S.recipe);
  const rec = R.reconcileColumns(S.recipe.columns, probe.available, includeNew);
  S.recipe.columns = rec.columns;
  if (!includeNew) S.columnNotes.added = rec.added;
}

function recompute() {
  const sh = sheet();
  if (!sh || !S.ranges.length) {
    S.result = null;
    S.markdown = '';
    return;
  }
  S.result = extract(sh, S.ranges, S.recipe, context());
  S.markdown = renderMarkdown(S.result, S.recipe, context());
}

// what: { source, user, side } — source: the selection, sheet or header row
// changed; user: the user did it (so a saved filter's source is modified);
// side: redraw the side panel (default yes).
function changed(what = {}) {
  if (what.user && what.source && S.active) S.active.sourceChanged = true;
  if (what.source) {
    syncSource(true);
    if (what.user) S.columnNotes.added = [];
  }
  recompute();
  if (what.side !== false) renderSide();
  renderToolbar();
  renderNotes();
  renderNameBox();
  renderDecor();
  renderPreview();
}

let pending = 0;
function changedSoon(what) {
  clearTimeout(pending);
  pending = setTimeout(() => changed(Object.assign({ side: false }, what)), 140);
}

function isModified() {
  if (!S.active) return false;
  return S.active.sourceChanged || R.isModified(S.recipe, S.active.baseline, false);
}

// --- selection ------------------------------------------------------------

function withEnd(range) {
  const sh = sheet();
  const clean = { top: range.top, left: range.left, bottom: range.bottom, right: range.right };
  clean.toEnd = describeRange(sh, clean, S.recipe.source.headerRow).toEnd;
  return clean;
}

function setRanges(ranges, active, user = true) {
  S.ranges = ranges.map(withEnd);
  S.activeRange = active === undefined ? S.ranges.length - 1 : active;
  grid.setSelection(S.ranges, S.activeRange);
  if (S.ranges.length) grid.scrollTo(S.ranges[S.activeRange]);
  changed({ source: true, user });
}

function onGridSelect(ranges, active, done) {
  if (!done) {
    S.ranges = ranges;
    S.activeRange = active;
    renderNameBox();
    return;
  }
  S.ranges = ranges.map((r) => Object.assign(withEnd(r), { anchor: r.anchor }));
  S.activeRange = active;
  changed({ source: true, user: true });
}

function onGridRegion(row, col, add) {
  const sh = sheet();
  if (!sh || row >= sh.rows || col >= sh.cols) return;
  const r = region.currentRegion(sh, row, col);
  if (add) setRanges(S.ranges.filter((x, i) => i !== S.activeRange).concat([r]));
  else setRanges([r], 0);
}

// --- saved filters --------------------------------------------------------

function onApplied(m) {
  const sh = S.data.get(m.sheetIndex);
  const recipe = R.normalize(m.recipe);
  const ask = R.askConditions(recipe);
  const commit = (values) => {
    for (const c of recipe.filter.conditions) {
      if (c.ask) c.values = values[c.id] || [];
    }
    S.sheetIndex = m.sheetIndex;
    grid.setSheet(sh);
    S.recipe = recipe;
    S.headerManual = true;
    S.ranges = m.ranges.map((r, i) => Object.assign({}, r, {
      toEnd: recipe.source.ranges[i] ? recipe.source.ranges[i].toEnd : false,
    }));
    S.activeRange = S.ranges.length - 1;
    grid.setSelection(S.ranges, S.activeRange);
    if (S.ranges.length) grid.scrollTo(S.ranges[0]);
    S.columnNotes = { added: [], missing: [] };
    syncSource(false);
    S.notes = m.notes || [];
    S.active = {
      id: m.filterId,
      name: m.name,
      baseline: clone(S.recipe),
      baseRanges: clone(S.ranges),
      baseSheet: S.sheetIndex,
      sourceChanged: false,
    };
    changed({});
    if (ask.length) post({ type: 'rememberValues', filterId: m.filterId, values });
  };
  if (!ask.length) {
    commit({});
    return;
  }
  askValues(m, recipe, ask, sh, commit);
}

// Asks only for the values the filter leaves open, with the values that
// exist in the data to choose from.
function askValues(m, recipe, ask, sh, commit) {
  const chosen = {};
  // The values on offer follow the fixed conditions and what is already
  // chosen in the other fields.
  const optionsFor = (cond) => {
    const probe = clone(recipe);
    for (const c of probe.filter.conditions) if (c.ask) c.values = chosen[c.id] || [];
    return distinctValues(sh, m.ranges, probe, cond.column, cond);
  };
  const fields = ask.map((cond) => {
    const op = F.operator(cond.op);
    const last = (m.lastValues && m.lastValues[cond.id]) || [];
    const known = new Set(optionsFor(cond).map((o) => o.value));
    chosen[cond.id] = op.multi && (cond.op === 'eq' || cond.op === 'neq') ? last.filter((v) => known.has(v)) : last.slice(0, op.multi ? undefined : 1);
    const label = refLabel(cond.column, T) + ' ' + T.t('op.' + cond.op);
    let control;
    if (op.multi) {
      control = valuePicker({
        values: chosen[cond.id],
        multi: true,
        placeholder: T.t('ui.any'),
        options: () => optionsFor(cond),
        onChange: (v) => { chosen[cond.id] = v; },
        t: T.t,
      });
    } else {
      control = h('input', {
        type: 'text',
        class: 'text',
        value: chosen[cond.id][0] || '',
        placeholder: T.t('ui.any'),
        onInput: (e) => { chosen[cond.id] = e.target.value.trim() ? [e.target.value] : []; },
      });
    }
    return h('div', { class: 'field' }, h('label', { class: 'field-label', text: label }), control);
  });
  const close = () => {
    closePopup();
    $.dialog.hidden = true;
    $.dialog.replaceChildren();
    document.removeEventListener('keydown', onKey, true);
  };
  const apply = () => {
    close();
    commit(chosen);
  };
  const cancel = () => {
    close();
    renderToolbar();
  };
  const onKey = (e) => {
    if (e.defaultPrevented) return;
    if (e.key === 'Escape' && !document.querySelector('.picker-pop')) {
      e.preventDefault();
      cancel();
    } else if (e.key === 'Enter' && !document.querySelector('.picker-pop') && e.target.tagName !== 'BUTTON') {
      e.preventDefault();
      apply();
    }
  };
  const box = h('div', { class: 'dialog', role: 'dialog', 'aria-modal': 'true' },
    h('h2', { text: T.t('ui.applyTitle', { name: m.name }) }),
    h('p', { class: 'muted', text: T.t('ui.applyHint') }),
    fields,
    h('div', { class: 'dialog-actions' },
      h('button', { class: 'primary', onClick: apply, text: T.t('ui.apply') }),
      h('button', { class: 'secondary', onClick: cancel, text: T.t('ui.cancel') })));
  $.dialog.replaceChildren(box);
  $.dialog.hidden = false;
  document.addEventListener('keydown', onKey, true);
  const first = box.querySelector('.picker, input');
  if (first) first.focus();
}

function recipeForSaving() {
  const sh = sheet();
  const r = clone(S.recipe);
  r.source.sheet = sh.name;
  r.source.ranges = S.ranges.map((rg) => Object.assign(describeRange(sh, rg, r.source.headerRow), { toEnd: !!rg.toEnd }));
  // Columns of the filter that this file does not have stay in the filter.
  if (S.active) {
    const saved = S.filters.find((f) => f.id === S.active.id);
    if (saved) {
      for (const c of saved.recipe.columns) {
        if (!r.columns.some((x) => refId(x) === refId(c))) r.columns.push(c);
      }
    }
  }
  return r;
}

function askedValues() {
  const values = {};
  for (const c of S.recipe.filter.conditions) if (c.ask && c.values.length) values[c.id] = c.values.slice();
  return values;
}

function saveAs() {
  if (!S.ranges.length) return;
  post({
    type: 'saveFilter',
    mode: 'new',
    recipe: recipeForSaving(),
    values: askedValues(),
    suggestedName: S.active ? S.active.name : '',
  });
}

function saveChanges() {
  if (!S.active) return;
  post({
    type: 'saveFilter',
    mode: 'update',
    filterId: S.active.id,
    keepSource: !S.active.sourceChanged,
    recipe: recipeForSaving(),
  });
}

function onSaved(m) {
  S.columnNotes.added = [];
  S.active = {
    id: m.filterId,
    name: m.name,
    baseline: clone(S.recipe),
    baseRanges: clone(S.ranges),
    baseSheet: S.sheetIndex,
    sourceChanged: false,
  };
  changed({ side: false });
}

async function revert() {
  if (!S.active) return;
  const a = S.active;
  const values = askedValues();
  if (a.baseSheet !== S.sheetIndex) {
    await loadSheet(a.baseSheet);
    S.sheetIndex = a.baseSheet;
    grid.setSheet(sheet());
  }
  S.recipe = clone(a.baseline);
  for (const c of S.recipe.filter.conditions) if (c.ask) c.values = values[c.id] || c.values;
  S.ranges = clone(a.baseRanges);
  S.activeRange = S.ranges.length - 1;
  S.headerManual = true;
  a.sourceChanged = false;
  grid.setSelection(S.ranges, S.activeRange);
  changed({});
}

// --- toolbar --------------------------------------------------------------

function renderToolbar() {
  const filters = S.filters;
  const matching = filters.filter((f) => f.matches);
  const other = filters.filter((f) => !f.matches);
  const options = [{ value: '', label: T.t('ui.noFilter') }];
  const item = (f) => ({ value: f.id, label: f.name });
  if (matching.length && other.length) {
    options.push({ group: T.t('ui.matchingFilters'), items: matching.map(item) });
    options.push({ group: T.t('ui.otherFilters'), items: other.map(item) });
  } else {
    for (const f of filters) options.push(item(f));
  }
  options.push({ value: '__manage__', label: T.t('ui.manageFilters') });
  const current = S.active ? S.active.id : '';
  const sel = select(options, current, (v) => {
    if (v === '__manage__') {
      sel.value = current;
      post({ type: 'manageFilters' });
    } else if (!v) {
      S.active = null;
      S.notes = [];
      changed({ side: false });
    } else {
      post({ type: 'applyFilter', filterId: v });
    }
  }, 'filter-select');
  sel.title = S.active ? (S.filters.find((f) => f.id === S.active.id) || {}).summary || '' : '';

  const modified = isModified();
  const hasData = !!S.ranges.length;
  const hasOutput = !!S.markdown;
  fill($.filterBar,
    h('label', { class: 'bar-label', text: T.t('ui.savedFilter') }),
    sel,
    modified ? h('span', { class: 'badge', title: T.t('ui.modifiedHint'), text: T.t('ui.modified') }) : null,
    modified ? h('button', { class: 'primary small', onClick: saveChanges, text: T.t('ui.saveChanges') }) : null,
    modified ? h('button', { class: 'link', onClick: revert, text: T.t('ui.revert') }) : null,
    h('button', {
      class: 'secondary small',
      disabled: !hasData,
      onClick: saveAs,
      text: S.active ? T.t('ui.saveAsNew') : T.t('ui.saveAsFilter'),
    }),
  );
  const nameHint = S.active ? S.active.name : '';
  fill($.exportBar,
    h('button', { class: 'secondary', disabled: !hasOutput, onClick: () => post({ type: 'copy', markdown: S.markdown }), text: T.t('ui.copy') }),
    h('button', { class: 'secondary', disabled: !hasOutput, onClick: () => post({ type: 'openUntitled', markdown: S.markdown }), text: T.t('ui.openInEditor') }),
    h('button', { class: 'primary', disabled: !hasOutput, onClick: () => post({ type: 'export', markdown: S.markdown, nameHint }), text: T.t('ui.export') }),
  );
}

function renderNotes() {
  const lines = [];
  for (const n of S.notes) lines.push(T.t('ui.note.' + n.code, n));
  if (S.result) {
    const miss = S.result.missing;
    if (miss.columns.length) lines.push(T.t('ui.columnsMissing', { names: miss.columns.map((c) => refLabel(c, T)).join(', ') }));
    for (const c of miss.conditions) lines.push(T.t('ui.conditionMissing', { name: refLabel(c.column, T) }));
  }
  if (S.columnNotes.added.length) {
    lines.push(T.t('ui.columnsNew', { names: S.columnNotes.added.map((c) => refLabel(c, T)).join(', ') }));
  }
  if (!lines.length) {
    $.notes.hidden = true;
    $.notes.replaceChildren();
    return;
  }
  $.notes.hidden = false;
  fill($.notes,
    h('div', { class: 'notes-text' }, lines.map((l) => h('div', { text: l }))),
    h('button', {
      class: 'icon',
      title: T.t('ui.cancel'),
      text: '×',
      onClick: () => {
        S.notes = [];
        S.columnNotes.added = [];
        $.notes.hidden = true;
      },
    }),
  );
}

// --- name box and sheet tabs ---------------------------------------------

function renderNameBox() {
  if (document.activeElement === $.nameBox) return;
  $.nameBox.value = S.ranges.map((r) => formatRange(r)).join(', ');
}

function onNameBoxKey(e) {
  if (e.key === 'Enter') {
    const sh = sheet();
    const ranges = sh ? parseRangeList($.nameBox.value, sh) : null;
    if (ranges) {
      $.nameBox.classList.remove('invalid');
      $.nameBox.blur();
      setRanges(ranges);
    } else {
      $.nameBox.classList.add('invalid');
    }
  } else if (e.key === 'Escape') {
    $.nameBox.classList.remove('invalid');
    $.nameBox.blur();
    renderNameBox();
  }
}

function renderTabs() {
  $.tabs.replaceChildren(...S.sheets.map((s, i) => h('button', {
    class: 'tab' + (i === S.sheetIndex ? ' is-active' : '') + (s.hidden ? ' is-hidden' : ''),
    title: s.hidden ? T.t('ui.hiddenSheet', { name: s.name }) : s.name,
    onClick: () => switchSheet(i).then(renderTabs),
    text: s.name,
  })));
}

function renderDecor() {
  const r = S.result;
  grid.setDecor({
    headerRows: new Set(S.recipe.source.headerRow ? S.ranges.map((x) => x.top) : []),
    dataRanges: S.ranges,
    keptRows: r ? r.keptRows : null,
    excludedCols: r ? r.excludedCols : new Set(),
  });
}

// --- side panel -----------------------------------------------------------

function section(key, title, badge, body) {
  const open = S.ui.sections[key] !== false;
  const head = h('button', {
    class: 'section-head',
    'aria-expanded': String(open),
    onClick: () => {
      S.ui.sections[key] = !open;
      saveUi();
      renderSide();
    },
  }, h('span', { class: 'twisty', text: open ? '▾' : '▸' }), h('span', { class: 'section-title', text: title }),
  badge ? h('span', { class: 'section-badge', text: badge }) : null);
  return h('section', { class: 'side-section' + (open ? ' is-open' : '') }, head, open ? h('div', { class: 'section-body' }, body) : null);
}

function renderSide() {
  const scroll = $.side.scrollTop;
  closePopup();
  fill($.side,
    sourceSection(),
    columnsSection(),
    conditionsSection(),
    optionsSection(),
  );
  $.side.scrollTop = scroll;
}

function sourceSection() {
  const sh = sheet();
  const body = [];
  if (!S.ranges.length) body.push(h('p', { class: 'muted', text: T.t('ui.noSelection') }));
  const list = h('div', { class: 'range-list' });
  S.ranges.forEach((r, i) => {
    const table = sh && (sh.tables || []).find((t) => t.range.top === r.top && t.range.left === r.left
      && t.range.bottom === r.bottom && t.range.right === r.right);
    list.appendChild(h('div', { class: 'range-row' + (i === S.activeRange ? ' is-active' : '') },
      h('button', {
        class: 'range-ref',
        title: formatRange(r),
        onClick: () => {
          S.activeRange = i;
          grid.setSelection(S.ranges, i);
          grid.scrollTo(r);
          renderSide();
        },
        text: table ? table.name + ' · ' + formatRange(r) : formatRange(r),
      }),
      checkbox(T.t('ui.toEnd'), r.toEnd, (v) => {
        S.ranges[i].toEnd = v;
        if (S.active) S.active.sourceChanged = true;
        changed({ side: false });
      }, T.t('ui.toEndHint')),
      h('button', {
        class: 'icon',
        title: T.t('ui.removeRange'),
        text: '×',
        onClick: () => setRanges(S.ranges.filter((x, j) => j !== i)),
      })));
  });
  if (S.ranges.length) body.push(list);

  const quick = h('div', { class: 'quick' });
  if (sh) {
    for (const t of sh.tables || []) {
      quick.appendChild(h('button', { class: 'pill', title: T.t('ui.excelTables') + ': ' + formatRange(t.range), onClick: () => setRanges([t.range], 0), text: t.name }));
    }
    const used = region.usedRange(sh);
    if (used) quick.appendChild(h('button', { class: 'pill', title: formatRange(used), onClick: () => setRanges([used], 0), text: T.t('ui.usedRange') }));
  }
  if (quick.childNodes.length) body.push(quick);

  if (S.ranges.length) {
    body.push(checkbox(T.t('ui.headerRow'), S.recipe.source.headerRow, (v) => {
      S.recipe.source.headerRow = v;
      S.headerManual = true;
      changed({ source: true, user: true });
    }));
  }
  if (S.ranges.length > 1) {
    body.push(h('div', { class: 'opt-row' }, h('span', { text: T.t('ui.multiRange') }),
      select(['auto', 'merge', 'separate'].map((v) => ({ value: v, label: T.t('ui.multiRange.' + v) })), S.recipe.options.multiRange, (v) => {
        S.recipe.options.multiRange = v;
        changed({});
      })));
  }
  body.push(h('p', { class: 'hint', text: T.t('ui.selectionHint') }));
  const badge = S.ranges.length > 1 ? S.ranges.length + ' ' + T.plural('plural.range', S.ranges.length) : '';
  return section('source', T.t('ui.section.source'), badge, body);
}

let dragIndex = -1;

function columnsSection() {
  const cols = S.recipe.columns;
  const body = [];
  if (!cols.length) {
    body.push(h('p', { class: 'muted', text: T.t('ui.noColumns') }));
    return section('columns', T.t('ui.section.columns'), '', body);
  }
  const included = cols.filter((c) => c.include).length;
  body.push(h('div', { class: 'col-tools' },
    h('button', { class: 'link', onClick: () => { cols.forEach((c) => { c.include = true; }); changed({}); }, text: T.t('ui.columnsAll') }),
    h('button', { class: 'link', onClick: () => { cols.forEach((c) => { c.include = false; }); changed({}); }, text: T.t('ui.columnsNone') }),
    h('span', { class: 'hint', text: T.t('ui.columnsHint') })));
  const types = new Map((S.result ? S.result.available : []).map((a) => [a.id, a.type]));
  const list = h('ul', { class: 'col-list' });
  cols.forEach((c, i) => {
    const input = h('input', {
      type: 'checkbox',
      checked: c.include,
      onChange: () => {
        c.include = input.checked;
        changed({});
      },
    });
    const li = h('li', {
      class: 'col-item' + (c.include ? '' : ' is-off'),
      draggable: 'true',
      onDragstart: (e) => {
        dragIndex = i;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', String(i));
        li.classList.add('is-dragging');
      },
      onDragend: () => li.classList.remove('is-dragging'),
      onDragover: (e) => {
        e.preventDefault();
        li.classList.add('is-drop');
      },
      onDragleave: () => li.classList.remove('is-drop'),
      onDrop: (e) => {
        e.preventDefault();
        li.classList.remove('is-drop');
        if (dragIndex < 0 || dragIndex === i) return;
        const [moved] = cols.splice(dragIndex, 1);
        cols.splice(i, 0, moved);
        dragIndex = -1;
        changed({});
      },
    },
    h('span', { class: 'grip', 'aria-hidden': 'true', text: '⋮⋮' }),
    h('label', { class: 'col-label' }, input, h('span', { class: 'col-name', text: refLabel(c, T) })),
    h('span', { class: 'col-meta', text: (types.get(refId(c)) === 'number' ? '#  ' : types.get(refId(c)) === 'date' ? '◷  ' : '') + c.letter }));
    list.appendChild(li);
  });
  body.push(list);

  const sortOptions = [{ value: '', label: T.t('ui.sortNone') }].concat(cols.map((c) => ({ value: refId(c), label: refLabel(c, T) })));
  const sortValue = S.recipe.sort.column ? refId(S.recipe.sort.column) : '';
  body.push(h('div', { class: 'opt-row' }, h('span', { text: T.t('ui.sortBy') }),
    select(sortOptions, sortValue, (v) => {
      const c = cols.find((x) => refId(x) === v);
      S.recipe.sort.column = c ? { name: c.name, letter: c.letter } : null;
      changed({});
    }),
    S.recipe.sort.column ? select([{ value: 'asc', label: T.t('ui.sortAsc') }, { value: 'desc', label: T.t('ui.sortDesc') }], S.recipe.sort.dir, (v) => {
      S.recipe.sort.dir = v;
      changed({});
    }) : null));
  return section('columns', T.t('ui.section.columns'), included + '/' + cols.length, body);
}

function conditionsSection() {
  const conds = S.recipe.filter.conditions;
  const available = S.result ? S.result.available : [];
  const body = [];
  if (conds.length > 1) {
    body.push(h('div', { class: 'opt-row' }, h('span', { text: T.t('ui.match') }),
      select([{ value: 'all', label: T.t('ui.match.all') }, { value: 'any', label: T.t('ui.match.any') }], S.recipe.filter.match, (v) => {
        S.recipe.filter.match = v;
        changed({});
      })));
  }
  if (!conds.length) body.push(h('p', { class: 'muted', text: T.t('ui.noConditions') }));
  conds.forEach((cond) => body.push(conditionCard(cond, available)));
  body.push(h('button', {
    class: 'secondary small add',
    disabled: !available.length,
    onClick: () => {
      const used = new Set(conds.map((c) => refId(c.column)));
      const first = available.find((a) => !used.has(a.id)) || available[0];
      conds.push({ id: R.newConditionId(), column: { name: first.name, letter: first.letter }, op: 'eq', values: [], ask: false });
      S.ui.sections.conditions = true;
      changed({});
    },
    text: '+ ' + T.t('ui.addCondition'),
  }));
  const active = conds.filter((c) => c.column && F.isActive(c, c.values)).length;
  return section('conditions', T.t('ui.section.conditions'), conds.length ? active + '/' + conds.length : '', body);
}

function conditionCard(cond, available) {
  const conds = S.recipe.filter.conditions;
  const col = findColumn(available, cond.column);
  const type = col ? col.type : 'text';
  const colOptions = available.map((a) => ({ value: a.id, label: refLabel(a, T) }));
  let colValue = col ? col.id : '__missing__';
  if (!col) colOptions.unshift({ value: '__missing__', label: refLabel(cond.column, T) + ' ⚠', disabled: true });
  const colSel = select(colOptions, colValue, (v) => {
    const a = available.find((x) => x.id === v);
    if (!a) return;
    cond.column = { name: a.name, letter: a.letter };
    const ops = F.operatorsFor(a.type);
    if (!ops.includes(cond.op)) cond.op = 'eq';
    cond.values = [];
    changed({});
  }, 'cond-col');
  const ops = F.operatorsFor(type);
  if (!ops.includes(cond.op)) ops.push(cond.op);
  const opSel = select(ops.map((o) => ({ value: o, label: T.t('op.' + o) })), cond.op, (v) => {
    const wasMulti = F.operator(cond.op).multi;
    cond.op = v;
    if (!F.operator(v).multi && wasMulti) cond.values = cond.values.slice(0, 1);
    if (F.operator(v).noValue) cond.ask = false;
    changed({});
  }, 'cond-op');
  const remove = h('button', {
    class: 'icon',
    title: T.t('ui.removeCondition'),
    text: '×',
    onClick: () => {
      S.recipe.filter.conditions = conds.filter((c) => c !== cond);
      changed({});
    },
  });
  const op = F.operator(cond.op);
  let valueRow = null;
  if (!op.noValue) {
    const placeholder = cond.ask ? T.t('ui.askPlaceholder')
      : type === 'number' && !op.multi ? T.t('ui.numberPlaceholder')
        : type === 'date' && !op.multi ? T.t('ui.datePlaceholder') : T.t('ui.choose');
    let control;
    if (op.multi) {
      control = valuePicker({
        values: cond.values,
        multi: true,
        placeholder,
        options: () => (sheet() && S.ranges.length ? distinctValues(sheet(), S.ranges, S.recipe, cond.column, cond) : []),
        onChange: (v) => {
          cond.values = v;
          changed({ side: false });
          updateSideBadges();
        },
        t: T.t,
      });
    } else {
      control = h('input', {
        type: 'text',
        class: 'text',
        value: cond.values[0] || '',
        placeholder,
        onInput: (e) => {
          cond.values = e.target.value.trim() ? [e.target.value] : [];
          changedSoon();
        },
      });
    }
    const mode = h('div', { class: 'seg', role: 'group' },
      h('button', {
        class: cond.ask ? '' : 'is-on',
        'aria-pressed': String(!cond.ask),
        title: T.t('ui.valueFixedHint'),
        onClick: () => { cond.ask = false; changed({}); },
        text: T.t('ui.valueFixed'),
      }),
      h('button', {
        class: cond.ask ? 'is-on' : '',
        'aria-pressed': String(cond.ask),
        title: T.t('ui.valueAskHint'),
        onClick: () => { cond.ask = true; changed({}); },
        text: T.t('ui.valueAsk'),
      }));
    valueRow = [h('div', { class: 'cond-value' }, control), mode];
  }
  return h('div', { class: 'cond' + (cond.ask ? ' is-ask' : '') + (col ? '' : ' is-missing') },
    h('div', { class: 'cond-head' }, colSel, opSel, remove),
    valueRow);
}

function updateSideBadges() {
  // Counts in section headers follow value edits without redrawing the panel.
  const conds = S.recipe.filter.conditions;
  const badge = $.side.querySelectorAll('.section-badge');
  const active = conds.filter((c) => c.column && F.isActive(c, c.values)).length;
  for (const b of badge) {
    const sectionTitle = b.parentElement.querySelector('.section-title');
    if (sectionTitle && sectionTitle.textContent === T.t('ui.section.conditions')) b.textContent = active + '/' + conds.length;
  }
}

function optionsSection() {
  const o = S.recipe.options;
  const set = (key) => (v) => {
    o[key] = v;
    changed({ source: key === 'skipHiddenColumns' || key === 'fillMerged' });
  };
  const choice = (key, values) => h('div', { class: 'opt-row' }, h('span', { text: T.t('ui.opt.' + key) }),
    select(values.map((v) => ({ value: v, label: T.t('ui.opt.' + key + '.' + v) })), o[key], set(key)));
  const body = [
    choice('layout', ['table', 'sections']),
    choice('title', ['none', 'sheet', 'filter', 'custom']),
    o.title === 'custom' ? h('div', { class: 'opt-row' }, h('span', { text: T.t('ui.opt.titleText') }),
      h('input', {
        type: 'text',
        class: 'text',
        value: o.titleText,
        onInput: (e) => {
          o.titleText = e.target.value;
          changedSoon();
        },
      })) : null,
    choice('values', ['formatted', 'raw']),
    choice('lineBreaks', ['br', 'space']),
    checkbox(T.t('ui.opt.skipHiddenRows'), o.skipHiddenRows, set('skipHiddenRows')),
    checkbox(T.t('ui.opt.skipHiddenColumns'), o.skipHiddenColumns, set('skipHiddenColumns')),
    checkbox(T.t('ui.opt.skipEmptyRows'), o.skipEmptyRows, set('skipEmptyRows')),
    checkbox(T.t('ui.opt.dropEmptyColumns'), o.dropEmptyColumns, set('dropEmptyColumns')),
    checkbox(T.t('ui.opt.fillMerged'), o.fillMerged, set('fillMerged')),
    o.layout === 'table' ? checkbox(T.t('ui.opt.alignNumbers'), o.alignNumbers, set('alignNumbers')) : null,
    o.layout === 'table' ? checkbox(T.t('ui.opt.padColumns'), o.padColumns, set('padColumns')) : null,
  ];
  return section('options', T.t('ui.section.options'), '', body);
}

// --- preview --------------------------------------------------------------

function truncated(result) {
  let left = PREVIEW_ROWS;
  let cut = false;
  const tables = result.tables.map((t) => {
    const rows = t.rows.slice(0, Math.max(0, left));
    if (rows.length < t.rows.length) cut = true;
    left -= rows.length;
    return Object.assign({}, t, { rows });
  });
  return { result: Object.assign({}, result, { tables }), cut };
}

function cellHtml(text, lineBreaks) {
  const parts = String(text).split(/\r?\n/);
  if (lineBreaks === 'space') return [parts.join(' ')];
  const out = [];
  parts.forEach((p, i) => {
    if (i) out.push(h('br'));
    out.push(p);
  });
  return out;
}

function renderedPreview(result, recipe) {
  const o = recipe.options;
  const ctx = context();
  const wrap = h('div', { class: 'md-rendered' });
  const titleText = o.title === 'sheet' ? ctx.sheetName : o.title === 'filter' ? ctx.filterName || ctx.sheetName
    : o.title === 'custom' ? o.titleText : '';
  if (titleText && titleText.trim()) wrap.appendChild(h('h2', { text: titleText.trim() }));
  const many = result.tables.length > 1;
  for (const t of result.tables) {
    if (!t.columns.length || (!t.rows.length && o.layout === 'sections')) continue;
    if (many) wrap.appendChild(h(titleText ? 'h3' : 'h2', { text: t.label }));
    if (o.layout === 'sections') {
      const level = Math.min(6, (titleText ? 3 : 2) + (many ? 1 : 0));
      for (const row of t.rows) {
        wrap.appendChild(h('h' + level, { text: row[0] || '—' }));
        const items = t.columns.map((c, i) => (i && row[i] ? h('li', null, h('strong', { text: (c.name || c.letter) + ':' }), ' ', ...cellHtml(row[i], o.lineBreaks)) : null)).filter(Boolean);
        if (items.length) wrap.appendChild(h('ul', null, items));
      }
    } else {
      const right = t.types.map((x) => o.alignNumbers && x === 'number');
      wrap.appendChild(h('table', null,
        h('thead', null, h('tr', null, t.columns.map((c, i) => h('th', { class: right[i] ? 'r' : '', text: c.name || c.letter })))),
        h('tbody', null, t.rows.map((row) => h('tr', null, row.map((v, i) => h('td', { class: right[i] ? 'r' : '' }, ...cellHtml(v, o.lineBreaks))))))));
    }
  }
  return wrap;
}

function renderPreview() {
  for (const b of $.previewModes.querySelectorAll('button')) {
    const on = b.dataset.mode === S.ui.preview;
    b.classList.toggle('is-on', on);
    b.setAttribute('aria-pressed', String(on));
  }
  const r = S.result;
  if (!r) {
    $.stats.textContent = '';
    $.previewBody.replaceChildren(h('p', { class: 'muted empty', text: T.t('ui.previewEmpty') }));
    return;
  }
  const cols = r.tables.reduce((n, t) => Math.max(n, t.columns.length), 0);
  $.stats.textContent = T.t('ui.stats', { kept: r.stats.kept, total: r.stats.rows, rows: T.plural('plural.row', r.stats.rows) })
    + ' · ' + T.t('ui.statsColumns', { count: cols, columns: T.plural('plural.column', cols) });
  const { result: shown, cut } = truncated(r);
  const body = [];
  if (!r.stats.kept) body.push(h('p', { class: 'muted', text: T.t('ui.previewNoRows') }));
  if (S.ui.preview === 'rendered') body.push(renderedPreview(shown, S.recipe));
  else body.push(h('pre', { class: 'md-code' }, h('code', { text: cut ? renderMarkdown(shown, S.recipe, context()) : S.markdown })));
  if (cut) body.push(h('p', { class: 'muted', text: T.t('ui.previewTruncated', { count: PREVIEW_ROWS }) }));
  fill($.previewBody, body);
}

// --- layout ---------------------------------------------------------------

function splitter(elm, axis, onMove) {
  elm.addEventListener('mousedown', (e) => {
    e.preventDefault();
    const start = axis === 'x' ? e.clientX : e.clientY;
    const move = (ev) => onMove((axis === 'x' ? ev.clientX : ev.clientY) - start, false);
    const up = (ev) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      onMove((axis === 'x' ? ev.clientX : ev.clientY) - start, true);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  });
}

function applyLayout() {
  document.documentElement.style.setProperty('--side-w', S.ui.sideWidth + 'px');
  document.documentElement.style.setProperty('--preview-h', S.ui.previewHeight + 'px');
}

function buildLayout() {
  const app = document.getElementById('app');
  app.className = 'app';
  $.filterBar = h('div', { class: 'filter-bar' });
  $.exportBar = h('div', { class: 'export-bar' });
  $.notes = h('div', { class: 'notes', role: 'status' });
  $.notes.hidden = true;
  $.nameBox = h('input', {
    class: 'name-box',
    type: 'text',
    spellcheck: 'false',
    'aria-label': T.t('ui.selection'),
    placeholder: T.t('ui.selectionPlaceholder'),
    onKeydown: onNameBoxKey,
  });
  $.gridHost = h('div', { class: 'grid' });
  $.tabs = h('nav', { class: 'sheet-tabs', 'aria-label': T.t('ui.sheet') });
  $.previewModes = h('div', { class: 'seg' },
    h('button', { 'data-mode': 'code', onClick: () => { S.ui.preview = 'code'; saveUi(); renderPreview(); }, text: T.t('ui.previewCode') }),
    h('button', { 'data-mode': 'rendered', onClick: () => { S.ui.preview = 'rendered'; saveUi(); renderPreview(); }, text: T.t('ui.previewRendered') }));
  $.stats = h('span', { class: 'stats' });
  $.previewBody = h('div', { class: 'preview-body' });
  $.side = h('aside', { class: 'side' });
  $.dialog = h('div', { class: 'dialog-backdrop' });
  $.dialog.hidden = true;
  const hsplit = h('div', { class: 'splitter h', role: 'separator', 'aria-orientation': 'horizontal' });
  const vsplit = h('div', { class: 'splitter v', role: 'separator', 'aria-orientation': 'vertical' });

  app.replaceChildren(
    h('header', { class: 'toolbar' },
      h('div', { class: 'file', title: S.fileName }, h('span', { class: 'file-icon', 'aria-hidden': 'true' }), h('span', { class: 'file-name', text: S.fileName })),
      $.filterBar,
      h('div', { class: 'spacer' }),
      $.exportBar),
    $.notes,
    h('main', { class: 'workspace' },
      h('section', { class: 'left' },
        h('div', { class: 'name-bar' }, $.nameBox),
        $.gridHost,
        $.tabs,
        hsplit,
        h('section', { class: 'preview' },
          h('div', { class: 'preview-head' }, h('span', { class: 'preview-title', text: T.t('ui.preview') }), $.previewModes, $.stats),
          $.previewBody)),
      vsplit,
      $.side),
    $.dialog);

  applyLayout();
  let baseSide = S.ui.sideWidth;
  splitter(vsplit, 'x', (d, end) => {
    S.ui.sideWidth = Math.max(280, Math.min(window.innerWidth - 360, baseSide - d));
    applyLayout();
    if (end) {
      baseSide = S.ui.sideWidth;
      saveUi();
    }
  });
  let basePreview = S.ui.previewHeight;
  splitter(hsplit, 'y', (d, end) => {
    S.ui.previewHeight = Math.max(80, Math.min(window.innerHeight - 220, basePreview - d));
    applyLayout();
    if (end) {
      basePreview = S.ui.previewHeight;
      saveUi();
    }
  });

  grid = new Grid($.gridHost, { onSelect: onGridSelect, onRegion: onGridRegion });
  grid.viewport.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      const used = sheet() && region.usedRange(sheet());
      if (used) setRanges([used], 0);
    }
  });
}

// --- messages -------------------------------------------------------------

function init(m) {
  T = fromMessages(m.messages, m.lang);
  S.fileName = m.fileName;
  S.sheets = m.sheets;
  S.sheetIndex = m.firstSheet;
  S.filters = m.filters;
  buildLayout();
  renderToolbar();
  renderSide();
  renderPreview();
  renderTabs();
}

function onSheet(m) {
  const sh = unpackSheet(m.sheet);
  S.data.set(m.index, sh);
  const waiters = S.waiting.get(m.index) || [];
  S.waiting.delete(m.index);
  for (const w of waiters) w(sh);
  if (m.index === S.sheetIndex && grid && grid.sheet !== sh && !grid.sheet) {
    grid.setSheet(sh);
    changed({});
  }
}

window.addEventListener('message', (e) => {
  const m = e.data;
  switch (m.type) {
    case 'init': init(m); break;
    case 'sheet': onSheet(m); break;
    case 'filters':
      S.filters = m.filters;
      if (S.active && !S.filters.some((f) => f.id === S.active.id)) S.active = null;
      else if (S.active) S.active.name = (S.filters.find((f) => f.id === S.active.id) || S.active).name;
      renderToolbar();
      break;
    case 'applied':
      onApplied(m);
      renderTabs();
      break;
    case 'saved': onSaved(m); break;
    case 'fatal': {
      const app = document.getElementById('app');
      app.className = 'fatal';
      app.textContent = m.message;
      if (m.action) {
        app.appendChild(h('div', { class: 'fatal-actions' },
          h('button', { class: 'primary', onClick: () => post({ type: 'openOriginal' }), text: m.action })));
      }
      break;
    }
    default: break;
  }
});

post({ type: 'ready' });
