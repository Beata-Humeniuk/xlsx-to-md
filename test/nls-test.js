'use strict';

const fs = require('fs');
const path = require('path');
const { check, eq } = require('./harness');
const { forLanguage, messagesFor, EN, TRANSLATIONS } = require('../src/nls');

check('every English key has a Polish translation', () => {
  const missing = Object.keys(EN).filter((k) => !(k in TRANSLATIONS.pl));
  eq(missing.join(', '), '', 'missing in pl');
  const extra = Object.keys(TRANSLATIONS.pl).filter((k) => !(k in EN));
  eq(extra.join(', '), '', 'unknown in pl');
});

check('placeholders match between languages', () => {
  for (const [key, en] of Object.entries(EN)) {
    if (typeof en !== 'string') continue;
    const names = (s) => (s.match(/\{\w+\}/g) || []).sort().join(',');
    eq(names(TRANSLATIONS.pl[key]), names(en), key);
  }
});

check('Polish plurals', () => {
  const pl = forLanguage('pl-PL');
  eq(pl.plural('plural.row', 1), 'wiersz');
  eq(pl.plural('plural.row', 3), 'wiersze');
  eq(pl.plural('plural.row', 12), 'wierszy');
  eq(pl.plural('plural.row', 22), 'wiersze');
  eq(pl.plural('plural.row', 25), 'wierszy');
});

check('unknown language falls back to English', () => {
  const de = forLanguage('de');
  eq(de.lang, 'en');
  eq(de.t('ui.apply'), 'Apply');
  eq(messagesFor('de')['ui.apply'], 'Apply');
});

check('parameters are filled in', () => {
  eq(forLanguage('pl').t('ui.applyTitle', { name: 'X' }), 'Filtr „X”');
});

check('package.nls files have the same keys', () => {
  const root = path.join(__dirname, '..');
  const en = JSON.parse(fs.readFileSync(path.join(root, 'package.nls.json'), 'utf8'));
  const pl = JSON.parse(fs.readFileSync(path.join(root, 'package.nls.pl.json'), 'utf8'));
  eq(Object.keys(pl).sort().join(','), Object.keys(en).sort().join(','));
  const manifest = fs.readFileSync(path.join(root, 'package.json'), 'utf8');
  for (const m of manifest.match(/%[\w.]+%/g)) {
    if (!(m.slice(1, -1) in en)) throw new Error('package.json uses ' + m + ' which package.nls.json lacks');
  }
});

require('./harness').done('nls');
