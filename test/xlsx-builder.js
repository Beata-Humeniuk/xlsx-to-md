'use strict';

// Builds .xlsx packages in memory for the tests: a tiny ZIP writer plus the
// minimum SpreadsheetML scaffolding around the given sheets.

const zlib = require('zlib');

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// files: { name: Buffer|string }; deflate: compress entries (method 8).
function zip(files, deflate = true) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, content] of Object.entries(files)) {
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const body = deflate ? zlib.deflateRawSync(data) : data;
    const nameBuf = Buffer.from(name, 'utf8');
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x800, 6);
    local.writeUInt16LE(deflate ? 8 : 0, 8);
    local.writeUInt32LE(crc32(data), 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x800, 8);
    central.writeUInt16LE(deflate ? 8 : 0, 10);
    central.writeUInt32LE(crc32(data), 16);
    central.writeUInt32LE(body.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, body);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + body.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(cd.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function letter(i) {
  let n = i + 1;
  let out = '';
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

// sheets: [{ name, rows: [[value]], hidden, hiddenRows: [index], hiddenCols:
// [index], merges: ['A1:B1'], tables: [{ name, ref }], styles: { 'B2': fmtIndex } }]
// Values: string → shared string, number → number, boolean → bool,
// { f, v } → formula with cached value, { d: serial, fmt } → date-formatted number.
// numFmts: extra custom formats [{ id, code }]; cellXfs reference them by index
// in `formats` (default list: General, date yyyy-mm-dd, 0.00, 0%).
function xlsx(sheets, options = {}) {
  const shared = [];
  const sharedIndex = new Map();
  const str = (s) => {
    if (!sharedIndex.has(s)) {
      sharedIndex.set(s, shared.length);
      shared.push(s);
    }
    return sharedIndex.get(s);
  };
  const formats = options.formats || [
    { id: 0 },
    { id: 164, code: 'yyyy-mm-dd' },
    { id: 2 },
    { id: 9 },
  ];
  const files = {};
  const sheetEntries = [];
  const wbRels = [];
  sheets.forEach((sheet, si) => {
    const n = si + 1;
    const rowsXml = sheet.rows.map((row, r) => {
      const cells = row.map((v, c) => {
        if (v === null || v === undefined || v === '') return '';
        const ref = letter(c) + (r + 1);
        const style = sheet.styles && sheet.styles[ref] !== undefined ? ' s="' + sheet.styles[ref] + '"' : '';
        if (typeof v === 'string') return '<c r="' + ref + '" t="s"' + style + '><v>' + str(v) + '</v></c>';
        if (typeof v === 'number') return '<c r="' + ref + '"' + style + '><v>' + v + '</v></c>';
        if (typeof v === 'boolean') return '<c r="' + ref + '" t="b"' + style + '><v>' + (v ? 1 : 0) + '</v></c>';
        if (v.d !== undefined) return '<c r="' + ref + '" s="' + (v.fmt === undefined ? 1 : v.fmt) + '"><v>' + v.d + '</v></c>';
        if (v.f !== undefined) {
          const t = typeof v.v === 'string' ? ' t="str"' : '';
          return '<c r="' + ref + '"' + t + style + '><f>' + esc(v.f) + '</f><v>' + esc(v.v) + '</v></c>';
        }
        if (v.inline !== undefined) return '<c r="' + ref + '" t="inlineStr"><is><t>' + esc(v.inline) + '</t></is></c>';
        return '';
      }).join('');
      const hidden = (sheet.hiddenRows || []).includes(r) ? ' hidden="1"' : '';
      return '<row r="' + (r + 1) + '"' + hidden + '>' + cells + '</row>';
    }).join('');
    const cols = (sheet.hiddenCols || []).map((c) => '<col min="' + (c + 1) + '" max="' + (c + 1) + '" hidden="1" width="0"/>').join('');
    const merges = (sheet.merges || []).map((m) => '<mergeCell ref="' + m + '"/>').join('');
    const sheetRels = [];
    const tableParts = (sheet.tables || []).map((t, ti) => {
      const id = 'rT' + (ti + 1);
      const file = 'table' + n + '_' + (ti + 1) + '.xml';
      sheetRels.push('<Relationship Id="' + id + '" Type="' + REL + '/table" Target="../tables/' + file + '"/>');
      files['xl/tables/' + file] = '<table xmlns="' + MAIN + '" id="' + (ti + 1) + '" name="' + t.name + '" displayName="' + t.name + '" ref="' + t.ref + '"/>';
      return '<tablePart r:id="' + id + '"/>';
    }).join('');
    files['xl/worksheets/sheet' + n + '.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<worksheet xmlns="' + MAIN + '" xmlns:r="' + REL + '">'
      + (cols ? '<cols>' + cols + '</cols>' : '')
      + '<sheetData>' + rowsXml + '</sheetData>'
      + (merges ? '<mergeCells count="' + sheet.merges.length + '">' + merges + '</mergeCells>' : '')
      + (tableParts ? '<tableParts count="' + sheet.tables.length + '">' + tableParts + '</tableParts>' : '')
      + '</worksheet>';
    if (sheetRels.length) {
      files['xl/worksheets/_rels/sheet' + n + '.xml.rels'] = '<Relationships xmlns="' + PKG + '">' + sheetRels.join('') + '</Relationships>';
    }
    sheetEntries.push('<sheet name="' + esc(sheet.name) + '" sheetId="' + n + '" r:id="rId' + n + '"' + (sheet.hidden ? ' state="hidden"' : '') + '/>');
    wbRels.push('<Relationship Id="rId' + n + '" Type="' + REL + '/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
  });
  wbRels.push('<Relationship Id="rIdS" Type="' + REL + '/sharedStrings" Target="sharedStrings.xml"/>');
  wbRels.push('<Relationship Id="rIdY" Type="' + REL + '/styles" Target="styles.xml"/>');
  files['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
  files['_rels/.rels'] = '<Relationships xmlns="' + PKG + '"><Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  files['xl/workbook.xml'] = '<workbook xmlns="' + MAIN + '" xmlns:r="' + REL + '">'
    + (options.date1904 ? '<workbookPr date1904="1"/>' : '')
    + '<sheets>' + sheetEntries.join('') + '</sheets></workbook>';
  files['xl/_rels/workbook.xml.rels'] = '<Relationships xmlns="' + PKG + '">' + wbRels.join('') + '</Relationships>';
  files['xl/sharedStrings.xml'] = '<sst xmlns="' + MAIN + '">' + shared.map((s) => '<si><t xml:space="preserve">' + esc(s) + '</t></si>').join('') + '</sst>';
  const custom = formats.filter((f) => f.code);
  files['xl/styles.xml'] = '<styleSheet xmlns="' + MAIN + '">'
    + (custom.length ? '<numFmts>' + custom.map((f) => '<numFmt numFmtId="' + f.id + '" formatCode="' + esc(f.code) + '"/>').join('') + '</numFmts>' : '')
    + '<cellXfs>' + formats.map((f) => '<xf numFmtId="' + f.id + '"/>').join('') + '</cellXfs></styleSheet>';
  return zip(files);
}

module.exports = { xlsx, zip };
