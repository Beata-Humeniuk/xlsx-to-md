'use strict';

// Reads an .xlsx/.xlsm package into plain sheet grids: every cell as the text
// Excel shows, its kind and (for numbers and dates) the raw number, plus the
// hidden rows and columns, merged areas and Excel tables of each sheet.
// Formulas are not evaluated; the value Excel cached when saving is used.

const path = require('path').posix;
const { readZip } = require('./zip');
const { parseXml, elements, child, childrenNamed, attr, textOf } = require('./xml');
const { formatNumber, builtinFormat, isDateFormat } = require('./core/numfmt');
const { parseCell, parseRange } = require('./core/address');

function sniff(buf) {
  if (!buf || buf.length < 8) return null;
  if (buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04) return 'xlsx';
  if (buf.readUInt32BE(0) === 0xd0cf11e0 && buf.readUInt32BE(4) === 0xa1b11ae1) return 'xls';
  return null;
}

function openPackage(buffer) {
  let zip;
  try {
    zip = readZip(buffer);
  } catch (e) {
    throw Object.assign(new Error('not an .xlsx package: ' + e.message), { code: 'NOT_XLSX' });
  }
  const names = new Map();
  for (const k of zip.keys()) names.set(k.toLowerCase(), k);
  const normalize = (p) => path.normalize(p.replace(/\\/g, '/')).replace(/^\/+/, '');
  const read = (p) => {
    const key = names.get(normalize(p).toLowerCase());
    return key ? zip.get(key)() : null;
  };
  const xml = (p) => {
    const buf = read(p);
    return buf ? parseXml(buf) : null;
  };
  // Relationships of a part; '' is the package itself (_rels/.rels).
  const rels = (partPath) => {
    const part = partPath ? normalize(partPath) : '';
    const dir = part ? path.dirname(part) : '';
    const relsPath = part ? (dir === '.' ? '' : dir + '/') + '_rels/' + path.basename(part) + '.rels' : '_rels/.rels';
    const map = new Map();
    const root = xml(relsPath);
    for (const r of childrenNamed(root, 'rel:Relationship').concat(childrenNamed(root, 'Relationship'))) {
      const target = attr(r, 'Target') || '';
      const external = attr(r, 'TargetMode') === 'External';
      map.set(attr(r, 'Id'), {
        type: attr(r, 'Type') || '',
        external,
        target: external ? target : target.startsWith('/') ? normalize(target) : normalize(path.join(dir, target)),
      });
    }
    return map;
  };
  let main = 'xl/workbook.xml';
  for (const r of rels('').values()) {
    if (r.type.endsWith('/officeDocument') && !r.external) main = r.target;
  }
  if (!read(main) || !/workbook/i.test(main)) {
    throw Object.assign(new Error('not an Excel workbook: ' + main + ' is missing'), { code: 'NOT_XLSX' });
  }
  return { read, xml, rels, main };
}

function readSharedStrings(root) {
  if (!root) return [];
  return childrenNamed(root, 's:si').map(richText);
}

// Text of an <si>/<is>: plain <t> or rich-text runs; phonetic runs (<rPh>) left out.
function richText(node) {
  let out = '';
  for (const c of elements(node)) {
    if (c.name === 's:t') out += textOf(c);
    else if (c.name === 's:r') out += textOf(child(c, 's:t'));
  }
  return out.replace(/_x([0-9A-Fa-f]{4})_/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function readStyles(root) {
  const custom = new Map();
  const numFmts = child(root, 's:numFmts');
  for (const f of childrenNamed(numFmts, 's:numFmt')) {
    custom.set(parseInt(attr(f, 'numFmtId'), 10), attr(f, 'formatCode') || '');
  }
  const xfs = childrenNamed(child(root, 's:cellXfs'), 's:xf').map((xf) => {
    const id = parseInt(attr(xf, 'numFmtId') || '0', 10);
    const code = custom.has(id) ? custom.get(id) : builtinFormat(id);
    return { code: code || 'General', date: isDateFormat(code) };
  });
  return xfs;
}

function sheetList(pkg, wb) {
  const relMap = pkg.rels(pkg.main);
  const sheets = [];
  for (const s of childrenNamed(child(wb, 's:sheets'), 's:sheet')) {
    const rel = relMap.get(attr(s, 'r:id'));
    if (!rel || rel.external) continue;
    // Chart sheets and dialog sheets carry no cells.
    if (!rel.type.endsWith('/worksheet')) continue;
    sheets.push({
      name: attr(s, 'name') || 'Sheet' + (sheets.length + 1),
      hidden: attr(s, 'state') === 'hidden' || attr(s, 'state') === 'veryHidden',
      part: rel.target,
    });
  }
  return sheets;
}

function readTables(pkg, sheetPart, sheetRoot) {
  const tables = [];
  const relMap = pkg.rels(sheetPart);
  for (const tp of childrenNamed(child(sheetRoot, 's:tableParts'), 's:tablePart')) {
    const rel = relMap.get(attr(tp, 'r:id'));
    if (!rel || rel.external) continue;
    const t = pkg.xml(rel.target);
    if (!t) continue;
    const range = parseRange(attr(t, 'ref') || '');
    if (!range) continue;
    tables.push({
      name: attr(t, 'displayName') || attr(t, 'name') || 'Table' + (tables.length + 1),
      range,
      headerRow: attr(t, 'headerRowCount') !== '0',
      totalsRow: parseInt(attr(t, 'totalsRowCount') || '0', 10) > 0,
    });
  }
  return tables;
}

function boolAttr(node, name) {
  const v = attr(node, name);
  return v === '1' || v === 'true';
}

// One sheet as { rows, cols, text[][], kind[][] ('s'|'n'|'d'|'b'|'e'|''),
// num[][] (number or null), hiddenRows[], hiddenCols[], merges[], tables[] }.
function readSheet(pkg, info, ctx) {
  const root = pkg.xml(info.part);
  if (!root) throw new Error('missing sheet part ' + info.part);
  const data = child(root, 's:sheetData');
  const cells = [];
  let maxRow = -1;
  let maxCol = -1;
  const hiddenRows = new Set();
  let nextRow = 0;
  for (const rowNode of childrenNamed(data, 's:row')) {
    const rAttr = attr(rowNode, 'r');
    const r = rAttr ? parseInt(rAttr, 10) - 1 : nextRow;
    nextRow = r + 1;
    if (boolAttr(rowNode, 'hidden')) hiddenRows.add(r);
    let nextCol = 0;
    for (const c of childrenNamed(rowNode, 's:c')) {
      const ref = attr(c, 'r');
      const pos = ref ? parseCell(ref) : { row: r, col: nextCol };
      const col = pos ? pos.col : nextCol;
      nextCol = col + 1;
      const cell = readCell(c, ctx);
      if (!cell) continue;
      if (!cells[r]) cells[r] = [];
      cells[r][col] = cell;
      if (r > maxRow) maxRow = r;
      if (col > maxCol) maxCol = col;
    }
  }

  const hiddenCols = new Set();
  for (const c of childrenNamed(child(root, 's:cols'), 's:col')) {
    if (!boolAttr(c, 'hidden')) continue;
    const min = parseInt(attr(c, 'min') || '1', 10) - 1;
    // Spans often run to the last sheet column; only the used width matters.
    const max = Math.min(parseInt(attr(c, 'max') || '1', 10) - 1, Math.max(maxCol, min));
    for (let i = min; i <= max; i++) hiddenCols.add(i);
  }

  const merges = [];
  for (const m of childrenNamed(child(root, 's:mergeCells'), 's:mergeCell')) {
    const range = parseRange(attr(m, 'ref') || '');
    if (range) merges.push(range);
  }

  const tables = readTables(pkg, info.part, root);
  for (const t of tables) {
    if (t.range.bottom > maxRow) maxRow = t.range.bottom;
    if (t.range.right > maxCol) maxCol = t.range.right;
  }

  const rows = maxRow + 1;
  const cols = maxCol + 1;
  const text = [];
  const kind = [];
  const num = [];
  for (let r = 0; r < rows; r++) {
    const src = cells[r] || [];
    const tRow = new Array(cols).fill('');
    const kRow = new Array(cols).fill('');
    const nRow = new Array(cols).fill(null);
    for (let c = 0; c < cols; c++) {
      const cell = src[c];
      if (!cell) continue;
      tRow[c] = cell.text;
      kRow[c] = cell.kind;
      nRow[c] = cell.num;
    }
    text.push(tRow);
    kind.push(kRow);
    num.push(nRow);
  }

  return {
    name: info.name,
    hidden: info.hidden,
    rows,
    cols,
    text,
    kind,
    num,
    hiddenRows: [...hiddenRows].filter((r) => r < rows).sort((a, b) => a - b),
    hiddenCols: [...hiddenCols].filter((c) => c < cols).sort((a, b) => a - b),
    merges,
    tables,
  };
}

function readCell(c, ctx) {
  const t = attr(c, 't') || 'n';
  const vNode = child(c, 's:v');
  const raw = vNode ? textOf(vNode) : null;
  const style = ctx.styles[parseInt(attr(c, 's') || '0', 10)] || ctx.styles[0] || { code: 'General', date: false };
  if (t === 'inlineStr') {
    const is = child(c, 's:is');
    const text = is ? richText(is) : raw || '';
    return text === '' ? null : { text, kind: 's', num: null };
  }
  if (raw === null || raw === '') return null;
  if (t === 's') {
    const text = ctx.shared[parseInt(raw, 10)];
    return text === undefined || text === '' ? null : { text, kind: 's', num: null };
  }
  if (t === 'str') return { text: raw, kind: 's', num: null };
  if (t === 'b') return { text: raw === '1' || raw === 'true' ? 'TRUE' : 'FALSE', kind: 'b', num: raw === '1' || raw === 'true' ? 1 : 0 };
  if (t === 'e') return { text: raw, kind: 'e', num: null };
  if (t === 'd') {
    // ISO 8601 date cell (Strict files): keep the date part, add time if any.
    const m = /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2})(:\d{2})?)?/.exec(raw);
    const text = m ? m[1] + (m[2] && m[2] !== '00:00' ? ' ' + m[2] : '') : raw;
    return { text, kind: 'd', num: isoToSerial(raw, ctx.date1904) };
  }
  const n = Number(raw);
  if (!Number.isFinite(n)) return { text: raw, kind: 's', num: null };
  return {
    text: formatNumber(n, style.code, ctx.date1904),
    kind: style.date ? 'd' : 'n',
    num: n,
  };
}

function isoToSerial(iso, date1904) {
  const ms = Date.parse(iso.length === 10 ? iso + 'T00:00:00Z' : /Z|[+-]\d\d:?\d\d$/.test(iso) ? iso : iso + 'Z');
  if (!Number.isFinite(ms)) return null;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (ms - epoch) / 86400000;
}

// Opens the workbook and returns { sheets: [{ name, hidden }], date1904,
// loadSheet(index) } — sheets are parsed on first use.
function openWorkbook(buffer) {
  const pkg = openPackage(buffer);
  const wb = pkg.xml(pkg.main);
  const relMap = pkg.rels(pkg.main);
  let sharedPart = null;
  let stylesPart = null;
  for (const r of relMap.values()) {
    if (r.type.endsWith('/sharedStrings')) sharedPart = r.target;
    if (r.type.endsWith('/styles')) stylesPart = r.target;
  }
  const props = child(wb, 's:workbookPr');
  const ctx = {
    shared: readSharedStrings(sharedPart ? pkg.xml(sharedPart) : null),
    styles: readStyles(stylesPart ? pkg.xml(stylesPart) : null),
    date1904: boolAttr(props, 'date1904'),
  };
  if (!ctx.styles.length) ctx.styles.push({ code: 'General', date: false });
  const list = sheetList(pkg, wb);
  if (!list.length) throw Object.assign(new Error('the workbook has no worksheets'), { code: 'NO_SHEETS' });
  const cache = new Map();
  return {
    sheets: list.map((s) => ({ name: s.name, hidden: s.hidden })),
    date1904: ctx.date1904,
    loadSheet(index) {
      if (!cache.has(index)) cache.set(index, readSheet(pkg, list[index], ctx));
      return cache.get(index);
    },
  };
}

module.exports = { openWorkbook, sniff };
