'use strict';

// Tables from extract() written as Markdown: GFM tables, or one section per
// row. Cell text is escaped only where Markdown would otherwise read it as
// syntax, so identifiers like field_name or List<String> stay readable.

function escapeCell(text, lineBreaks) {
  let s = String(text == null ? '' : text).replace(/\r\n?/g, '\n').trim();
  s = s.replace(/\\(?=[\\`*_{}[\]()#+\-.!|<>~])/g, '\\\\');
  s = s.replace(/\|/g, '\\|');
  // Something that would be read as an HTML tag or an autolink.
  s = s.replace(/<(?=[A-Za-z/!?])/g, '\\<');
  // Emphasis markers at word edges; intraword underscores are harmless in GFM.
  s = s.replace(/\*/g, '\\*');
  s = s.replace(/(^|[^\p{L}\p{N}_])_|_(?=$|[^\p{L}\p{N}_])/gu, (m, before) => (before !== undefined ? before + '\\_' : '\\_'));
  s = s.replace(/~~/g, '\\~\\~');
  s = s.replace(/\n+/g, lineBreaks === 'space' ? ' ' : '<br>');
  return s;
}

// Text at the start of a line that Markdown would turn into a heading, list
// or quote when it is not inside a table.
function escapeBlock(text, lineBreaks) {
  let s = escapeCell(text, lineBreaks).replace(/\\\|/g, '|');
  s = s.replace(/^(#{1,6}\s|>|[-+]\s|\d+[.)]\s)/, '\\$1');
  return s;
}

function width(s) {
  let n = 0;
  for (const ch of s) {
    const code = ch.codePointAt(0);
    // East Asian wide characters take two columns in a monospaced editor.
    n += (code >= 0x1100 && code <= 0x115f) || (code >= 0x2e80 && code <= 0xa4cf)
      || (code >= 0xac00 && code <= 0xd7a3) || (code >= 0xf900 && code <= 0xfaff)
      || (code >= 0xfe30 && code <= 0xfe4f) || (code >= 0xff00 && code <= 0xff60)
      || (code >= 0xffe0 && code <= 0xffe6) || (code >= 0x1f300 && code <= 0x1faff) ? 2 : 1;
  }
  return n;
}

const PAD_LIMIT = 60;

function gfmTable(table, options) {
  const headers = table.columns.map((c) => escapeCell(c.name || c.letter, options.lineBreaks));
  const body = table.rows.map((row) => row.map((v) => escapeCell(v, options.lineBreaks)));
  const right = table.types.map((t) => options.alignNumbers && t === 'number');
  const all = [headers, ...body];
  const widths = headers.map((_, i) => Math.max(3, ...all.map((r) => width(r[i] || ''))));
  const pad = options.padColumns && widths.every((w) => w <= PAD_LIMIT);
  const fit = (s, i) => {
    if (!pad) return s;
    const gap = ' '.repeat(Math.max(0, widths[i] - width(s)));
    return right[i] ? gap + s : s + gap;
  };
  const line = (cells) => '| ' + cells.map((s, i) => fit(s || '', i)).join(' | ') + ' |';
  const rule = '|' + widths.map((w, i) => {
    const n = pad ? w : 3;
    return right[i] ? ' ' + '-'.repeat(n - 1) + ': ' : ' ' + '-'.repeat(n) + ' ';
  }).join('|') + '|';
  return [line(headers), rule, ...body.map(line)].join('\n');
}

// One section per row: the first column is the heading, the others a list
// of "**Header:** value" lines (empty values left out).
function sections(table, options, level) {
  const hashes = '#'.repeat(level);
  const out = [];
  for (const row of table.rows) {
    const title = escapeBlock(row[0] || '', 'space') || '—';
    const lines = [hashes + ' ' + title, ''];
    table.columns.forEach((c, i) => {
      if (i === 0 || !row[i]) return;
      const name = escapeBlock(c.name || c.letter, 'space');
      lines.push('- **' + name + ':** ' + escapeCell(row[i], options.lineBreaks).replace(/\\\|/g, '|'));
    });
    if (lines.length === 2) lines.pop();
    out.push(lines.join('\n'));
  }
  return out.join('\n\n');
}

function titleText(recipe, context) {
  const o = recipe.options;
  if (o.title === 'sheet') return context.sheetName || '';
  if (o.title === 'filter') return context.filterName || context.sheetName || '';
  if (o.title === 'custom') return o.titleText || '';
  return '';
}

// result: from extract(); recipe: normalized; context: { sheetName, filterName }.
function renderMarkdown(result, recipe, context = {}) {
  const o = recipe.options;
  const title = titleText(recipe, context).trim();
  const blocks = [];
  if (title) blocks.push('## ' + escapeBlock(title, 'space'));
  const many = result.tables.length > 1;
  for (const table of result.tables) {
    if (!table.columns.length) continue;
    // A section per row has nothing to show for a table without rows.
    if (!table.rows.length && o.layout === 'sections') continue;
    if (many) blocks.push((title ? '### ' : '## ') + escapeBlock(table.label, 'space'));
    if (!table.rows.length) {
      blocks.push(gfmTable(table, o));
      continue;
    }
    const level = (title ? 3 : 2) + (many ? 1 : 0);
    blocks.push(o.layout === 'sections' ? sections(table, o, level) : gfmTable(table, o));
  }
  return blocks.length ? blocks.join('\n\n') + '\n' : '';
}

module.exports = { renderMarkdown, escapeCell, gfmTable };
