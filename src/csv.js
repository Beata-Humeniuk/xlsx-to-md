'use strict';

// Reads a CSV/TSV file into the same sheet shape as the workbook reader, so
// the editor, conditions and saved filters work on it unchanged. The
// encoding (UTF-8, UTF-16 with BOM, or Windows-1250 as Excel writes it in
// Polish Windows — even mixed line by line) and the separator (, ; tab |)
// are recognized from the content; Excel's "sep=;" first line is honoured.

const { parseNumber, parseDate } = require('./core/filter');
const { decodeText } = require('./text');

const SEPARATORS = [',', ';', '\t', '|'];

function decode(buf) {
  return decodeText(buf);
}

// Separator counts per line outside quotes, for the first lines.
function countSeparators(text, sep, maxLines) {
  const counts = [];
  let n = 0;
  let quoted = false;
  for (let i = 0; i < text.length && counts.length < maxLines; i++) {
    const ch = text[i];
    if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === sep) n++;
    else if (!quoted && ch === '\n') {
      counts.push(n);
      n = 0;
    }
  }
  if (counts.length < maxLines && n) counts.push(n);
  return counts;
}

// The separator that splits the first lines most consistently.
function detectSeparator(text) {
  let best = ',';
  let bestScore = -1;
  for (const sep of SEPARATORS) {
    const counts = countSeparators(text, sep, 30);
    if (!counts.length || !counts[0]) continue;
    const first = counts[0];
    const same = counts.filter((c) => c === first).length;
    const score = same * 1000 + first;
    if (score > bestScore) {
      best = sep;
      bestScore = score;
    }
  }
  return best;
}

// RFC 4180 rows: quoted fields with "" for a quote and line breaks inside.
function parseRows(text, sep) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
      } else {
        field += ch;
      }
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
    } else if (ch === sep) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      if (ch === '\r' && text[i + 1] === '\n') i++;
    } else {
      field += ch;
    }
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

// A CSV cell has no type; numbers and dates are recognized from the text so
// that conditions compare them as such. The text itself stays as written.
function classify(text) {
  const t = text.trim();
  if (t === '') return { kind: '', num: null };
  if (/^(true|false)$/i.test(t)) return { kind: 'b', num: /^true$/i.test(t) ? 1 : 0 };
  // Codes with leading zeros ("007") stay text.
  if (/^[-+]?[\d\s.,]*\d%?$/.test(t) && !/^[-+]?0\d/.test(t)) {
    const n = parseNumber(t);
    if (n !== null) return { kind: 'n', num: n };
  }
  const d = parseDate(t, false);
  if (d !== null) return { kind: 'd', num: d };
  return { kind: 's', num: null };
}

function readCsv(buffer, name, options = {}) {
  let text = decode(buffer).replace(/^﻿/, '');
  let sep = options.separator || null;
  const sepLine = /^sep=(.)\r?\n/i.exec(text);
  if (sepLine) {
    sep = sep || sepLine[1];
    text = text.slice(sepLine[0].length);
  }
  if (!sep) sep = options.tsv ? '\t' : detectSeparator(text);
  const raw = parseRows(text, sep);
  while (raw.length && raw[raw.length - 1].every((f) => f === '')) raw.pop();
  const cols = raw.reduce((n, r) => Math.max(n, r.length), 0);
  const textRows = [];
  const kindRows = [];
  const numRows = [];
  for (const r of raw) {
    const tRow = new Array(cols).fill('');
    const kRow = new Array(cols).fill('');
    const nRow = new Array(cols).fill(null);
    r.forEach((v, c) => {
      tRow[c] = v;
      const k = classify(v);
      kRow[c] = k.kind;
      nRow[c] = k.num;
    });
    textRows.push(tRow);
    kindRows.push(kRow);
    numRows.push(nRow);
  }
  return {
    name,
    hidden: false,
    rows: raw.length,
    cols,
    text: textRows,
    kind: kindRows,
    num: numRows,
    hiddenRows: [],
    hiddenCols: [],
    merges: [],
    tables: [],
    separator: sep,
  };
}

// The same interface as openWorkbook: one sheet named after the file.
function openCsv(buffer, fileName) {
  const name = String(fileName).replace(/\.[^.]+$/, '') || 'CSV';
  const sheet = readCsv(buffer, name, { tsv: /\.tsv$/i.test(fileName) });
  return {
    sheets: [{ name, hidden: false }],
    date1904: false,
    loadSheet() {
      return sheet;
    },
  };
}

function isCsvName(file) {
  return /\.(csv|tsv)$/i.test(String(file));
}

module.exports = { openCsv, readCsv, detectSeparator, parseRows, decode, isCsvName };
