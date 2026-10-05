'use strict';

// Column identity. A column is known by its header text first — Excel column
// order changes from file to file, names much less often — and by its letter
// only when there is no header to go by.

const { columnLetter, columnIndex } = require('./address');

// Header text as written in the output: one line, single spaces.
function cleanHeader(text) {
  return String(text == null ? '' : text).replace(/\s+/g, ' ').trim();
}

// Header text for matching: case, accents of the Unicode form and extra
// spaces do not matter, nor a trailing colon or asterisk ("Required*").
function headerKey(text) {
  return cleanHeader(text).normalize('NFKC').toLowerCase().replace(/\s*[:*]+$/, '');
}

// A column reference as stored in recipes: { name, letter }. `name` is the
// header text (null without a header row), `letter` the sheet column letter
// it had when the reference was made — the fallback.
function columnRef(name, col) {
  return { name: name ? cleanHeader(name) : null, letter: columnLetter(col) };
}

function refId(ref) {
  if (!ref) return '';
  return ref.name ? 'h:' + headerKey(ref.name) : 'c:' + String(ref.letter || '').toUpperCase();
}

function refLabel(ref, nls) {
  if (!ref) return '';
  if (ref.name) return ref.name;
  return nls ? nls.t('column.letter', { letter: ref.letter }) : 'Column ' + ref.letter;
}

// The columns of a range from [{ text, col }] (header text and sheet column
// index): the id used to match recipes, the display name and the letter.
// Duplicate header names get " (2)", " (3)" so each stays addressable.
function rangeColumns(entries, hasHeader) {
  const seen = new Map();
  return entries.map(({ text, col }) => {
    const clean = hasHeader ? cleanHeader(text) : '';
    if (!clean) {
      return { id: 'c:' + columnLetter(col), name: null, letter: columnLetter(col), col };
    }
    const key = headerKey(clean);
    const n = (seen.get(key) || 0) + 1;
    seen.set(key, n);
    const name = n === 1 ? clean : clean + ' (' + n + ')';
    return { id: 'h:' + headerKey(name), name, letter: columnLetter(col), col };
  });
}

// Finds the column a reference points at among `columns` (from rangeColumns):
// by name when the reference has one, by letter otherwise or when the data
// has no names at all.
function findColumn(columns, ref) {
  if (!ref) return null;
  if (ref.name) {
    const id = 'h:' + headerKey(ref.name);
    const hit = columns.find((c) => c.id === id);
    if (hit) return hit;
    if (columns.some((c) => c.name)) return null;
  }
  const letter = String(ref.letter || '').toUpperCase();
  return columns.find((c) => c.letter === letter) || null;
}

module.exports = { cleanHeader, headerKey, columnRef, refId, refLabel, rangeColumns, findColumn, columnIndex };
