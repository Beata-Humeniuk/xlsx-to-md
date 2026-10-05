'use strict';

const { check, eq } = require('./harness');
const { openCsv, readCsv, detectSeparator, parseRows, decode } = require('../src/csv');
const { extract } = require('../src/core/extract');
const { renderMarkdown } = require('../src/core/markdown');
const { resolveSource, describeRange } = require('../src/core/resolve');
const R = require('../src/core/recipe');

check('separators', () => {
  eq(detectSeparator('a,b,c\n1,2,3\n'), ',');
  eq(detectSeparator('a;b;c\n1,5;2,5;3\n'), ';');
  eq(detectSeparator('a\tb\n1\t2\n'), '\t');
  eq(detectSeparator('a|b\n1|2\n'), '|');
  eq(detectSeparator('"x,y";b\n"1,2";3\n'), ';');
});

check('quotes, doubled quotes, line breaks, CRLF', () => {
  const rows = parseRows('a,"b ""q"", c","line\r\nbreak"\r\n1,2,3\r\n', ',');
  eq(JSON.stringify(rows), JSON.stringify([['a', 'b "q", c', 'line\r\nbreak'], ['1', '2', '3']]));
});

check('encodings', () => {
  eq(decode(Buffer.from('﻿Zażółć', 'utf8')), 'Zażółć');
  eq(decode(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('Łódź', 'utf16le')])), 'Łódź');
  // "Łódź" in Windows-1250, as Excel saves CSV on Polish Windows.
  eq(decode(Buffer.from([0xa3, 0xf3, 0x64, 0x9f])), 'Łódź');
});

check('mixed file: UTF-8 lines and Windows-1250 lines', () => {
  const utf = Buffer.from('Kod;Opis\n1;Błąd zwracany\n', 'utf8');
  const cp = Buffer.from([0x32, 0x3b, 0x42, 0xb3, 0xb9, 0x64, 0x0a]); // "2;Błąd\n" in Windows-1250
  const s = readCsv(Buffer.concat([utf, cp]), 'x');
  eq(s.text.map((r) => r[1]).join('|'), 'Opis|Błąd zwracany|Błąd');
});

check('sep= line, ragged rows, trailing empty lines', () => {
  const s = readCsv(Buffer.from('sep=;\nA;B;C\n1;2\n\n\n'), 'x');
  eq(s.separator, ';');
  eq(s.rows, 2);
  eq(s.cols, 3);
  eq(s.text[1].join('|'), '1|2|');
});

check('types recognized, text kept as written', () => {
  const s = readCsv(Buffer.from('n;d;b;code;t\n1 234,50;2026-03-15;TRUE;007;abc\n'), 'x');
  eq(s.kind[1].join(''), 'ndbss');
  eq(s.num[1][0], 1234.5);
  eq(s.text[1][0], '1 234,50');
  eq(s.text[1][3], '007');
});

const CSV = 'Field;Type;Status;System\nid;string;OK;CRM\nold;int;DELETED;CRM\namount;decimal;OK;FUND\n';

check('same pipeline as workbooks', () => {
  const wb = openCsv(Buffer.from(CSV), 'spec.csv');
  eq(wb.sheets[0].name, 'spec');
  const sheet = wb.loadSheet(0);
  const r = R.emptyRecipe();
  r.columns = [{ name: 'Field', letter: 'A', include: true }, { name: 'System', letter: 'D', include: true }];
  r.filter.conditions = [{ id: 'a', column: { name: 'Status' }, op: 'neq', values: ['DELETED'], ask: false }];
  const n = R.normalize(r);
  const md = renderMarkdown(extract(sheet, [{ top: 0, left: 0, bottom: 3, right: 3 }], n), n);
  eq(md, '| Field  | System |\n| ------ | ------ |\n| id     | CRM    |\n| amount | FUND   |\n');
});

check('a filter saved on one CSV finds its data in another', () => {
  const first = openCsv(Buffer.from(CSV), 'spec-v1.csv');
  const r = R.emptyRecipe();
  r.source.sheet = 'spec-v1';
  r.source.ranges = [describeRange(first.loadSheet(0), { top: 0, left: 0, bottom: 3, right: 3 }, true)];
  const saved = R.forSaving(r);
  const second = openCsv(Buffer.from('Owner,Field,Type,Status,System\na,id,string,OK,CRM\nb,x,int,OK,ESB\nc,y,int,OK,ESB\nd,z,int,OK,ESB\n'), 'spec-v2.csv');
  const hit = resolveSource(second, saved);
  eq(hit.notes.length, 0);
  eq(JSON.stringify(hit.ranges), JSON.stringify([{ top: 0, left: 0, bottom: 4, right: 4 }]));
});

require('./harness').done('csv');
