'use strict';

const { check, eq } = require('./harness');
const { xlsx, zip } = require('./xlsx-builder');
const { openWorkbook, sniff } = require('../src/workbook');

const book = xlsx([
  {
    name: 'Data',
    rows: [
      ['Title spanning'],
      ['Name', 'Count', 'Ok', 'When', 'Formula', 'Inline'],
      ['alpha', 3, true, { d: 45000 }, { f: 'B3*2', v: 6 }, { inline: 'in' }],
      ['beta', 2.5, false, { d: 45001.25, fmt: 0 }, { f: 'CONCAT(A4)', v: 'beta!' }, null],
    ],
    hiddenRows: [3],
    hiddenCols: [5],
    merges: ['A1:C1'],
    tables: [{ name: 'Items', ref: 'A2:F4' }],
    styles: { B4: 2 },
  },
  { name: 'Secret', hidden: true, rows: [['x']] },
]);

check('sniff', () => {
  eq(sniff(book), 'xlsx');
  eq(sniff(Buffer.from('d0cf11e0a1b11ae1000000', 'hex')), 'xls');
  eq(sniff(Buffer.from('hello world')), null);
});

const wb = openWorkbook(book);

check('sheet list with hidden state', () => {
  eq(JSON.stringify(wb.sheets), JSON.stringify([{ name: 'Data', hidden: false }, { name: 'Secret', hidden: true }]));
});

const s = wb.loadSheet(0);

check('size', () => {
  eq(s.rows, 4);
  eq(s.cols, 6);
});

check('cell text and kinds', () => {
  eq(s.text[1].join('|'), 'Name|Count|Ok|When|Formula|Inline');
  eq(s.text[2].join('|'), 'alpha|3|TRUE|2023-03-15|6|in');
  eq(s.kind[2].join(''), 'snbdns');
  eq(s.num[2][1], 3);
  eq(s.num[2][3], 45000);
});

check('number formats and formula strings', () => {
  eq(s.text[3][1], '2.50');
  eq(s.text[3][4], 'beta!');
  eq(s.kind[3][4], 's');
  eq(s.text[3][3], '45001.25');
});

check('hidden rows, columns, merges, tables', () => {
  eq(s.hiddenRows.join(','), '3');
  eq(s.hiddenCols.join(','), '5');
  eq(JSON.stringify(s.merges), JSON.stringify([{ top: 0, left: 0, bottom: 0, right: 2 }]));
  eq(s.tables.length, 1);
  eq(s.tables[0].name, 'Items');
  eq(JSON.stringify(s.tables[0].range), JSON.stringify({ top: 1, left: 0, bottom: 3, right: 5 }));
});

check('not a workbook', () => {
  let code = null;
  try {
    openWorkbook(zip({ 'word/document.xml': '<w:document/>', '_rels/.rels': '<Relationships/>' }));
  } catch (e) {
    code = e.code;
  }
  eq(code, 'NOT_XLSX');
});

check('rich text and escaped characters in shared strings', () => {
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const buf = zip({
    '_rels/.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="r" Type="' + REL + '/officeDocument" Target="/xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<x:workbook xmlns:x="' + MAIN + '" xmlns:r="' + REL + '"><x:sheets><x:sheet name="S" sheetId="1" r:id="a"/></x:sheets></x:workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="a" Type="' + REL + '/worksheet" Target="s.xml"/><Relationship Id="b" Type="' + REL + '/sharedStrings" Target="ss.xml"/></Relationships>',
    'xl/ss.xml': '<sst xmlns="' + MAIN + '"><si><r><t>Bold</t></r><r><t xml:space="preserve"> part</t></r><rPh><t>x</t></rPh></si><si><t>Line_x000A_two</t></si></sst>',
    'xl/s.xml': '<worksheet xmlns="' + MAIN + '"><sheetData><row><c t="s"><v>0</v></c><c t="s"><v>1</v></c></row></sheetData></worksheet>',
  });
  const one = openWorkbook(buf).loadSheet(0);
  eq(one.text[0][0], 'Bold part');
  eq(one.text[0][1], 'Line\ntwo');
});

check('parts in a declared encoding (windows-1250)', () => {
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const cp1250 = (s) => Buffer.from(Array.from(s, (ch) => ({ 'ł': 0xb3, 'ą': 0xb9, 'ż': 0xbf, 'ó': 0xf3 }[ch] || ch.charCodeAt(0))));
  const buf = zip({
    '_rels/.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="r" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="' + MAIN + '" xmlns:r="' + REL + '"><sheets><sheet name="S" sheetId="1" r:id="w"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="w" Type="' + REL + '/worksheet" Target="s.xml"/>'
      + '<Relationship Id="t" Type="' + REL + '/sharedStrings" Target="ss.xml"/></Relationships>',
    'xl/ss.xml': cp1250('<?xml version="1.0" encoding="windows-1250"?><sst xmlns="' + MAIN + '"><si><t>Błąd bieżącego</t></si></sst>'),
    'xl/s.xml': cp1250('<?xml version="1.0" encoding="Windows-1250"?><worksheet xmlns="' + MAIN + '"><sheetData><row r="1">'
      + '<c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Usługa ó</t></is></c></row></sheetData></worksheet>'),
  });
  const one = openWorkbook(buf).loadSheet(0);
  eq(one.text[0].join('|'), 'Błąd bieżącego|Usługa ó');
});

check('UTF-8 declared but not: read with the code page', () => {
  const { decodeXml } = require('../src/text');
  eq(decodeXml(Buffer.concat([Buffer.from('<?xml version="1.0"?><t>B'), Buffer.from([0xb3, 0xb9]), Buffer.from('d</t>')])), '<?xml version="1.0"?><t>Błąd</t>');
  eq(decodeXml(Buffer.from('<t>Błąd</t>', 'utf8')), '<t>Błąd</t>');
});

check('macro-enabled workbook (.xlsm): macros and macro sheets are skipped', () => {
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const buf = zip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="r" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="' + MAIN + '" xmlns:r="' + REL + '"><sheets>'
      + '<sheet name="Macro1" sheetId="1" r:id="m"/><sheet name="Data" sheetId="2" r:id="w"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="' + PKG + '">'
      + '<Relationship Id="m" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>'
      + '<Relationship Id="w" Type="' + REL + '/worksheet" Target="worksheets/sheet1.xml"/>'
      + '<Relationship Id="v" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    'xl/macrosheets/sheet1.xml': '<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/>',
    'xl/vbaProject.bin': Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
    'xl/worksheets/sheet1.xml': '<worksheet xmlns="' + MAIN + '"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Field</t></is></c></row>'
      + '<row r="2"><c r="A2" t="inlineStr"><is><t>id</t></is></c></row></sheetData></worksheet>',
  });
  eq(sniff(buf), 'xlsx');
  const wb = openWorkbook(buf);
  eq(JSON.stringify(wb.sheets), JSON.stringify([{ name: 'Data', hidden: false }]));
  eq(wb.loadSheet(0).text.map((r) => r[0]).join(','), 'Field,id');
});

require('./harness').done('workbook');
