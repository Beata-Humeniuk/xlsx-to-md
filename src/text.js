'use strict';

// Bytes to text the way Excel reads them. XML parts follow their BOM or
// their declared encoding (some exporters write encoding="windows-1250");
// text files without a BOM are UTF-8 where they are valid UTF-8 and the
// Windows code page (1250, Central European) elsewhere — line by line, as
// logs appended by different programs mix both.

const FALLBACK = 'windows-1250';

function decoder(label, fatal) {
  try {
    return new TextDecoder(label, { fatal: !!fatal });
  } catch {
    return null;
  }
}

function bom(buf) {
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return { encoding: 'utf-8', skip: 3 };
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return { encoding: 'utf-16le', skip: 2 };
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) return { encoding: 'utf-16be', skip: 2 };
  return null;
}

function withLabel(buf, label) {
  const d = decoder(label, false) || decoder(FALLBACK, false);
  return d ? d.decode(buf) : buf.toString('latin1');
}

function isUtf8(buf) {
  try {
    decoder('utf-8', true).decode(buf);
    return true;
  } catch {
    return false;
  }
}

function decodeXml(buf) {
  if (!Buffer.isBuffer(buf)) return String(buf);
  const b = bom(buf);
  if (b) return withLabel(buf.subarray(b.skip), b.encoding);
  const head = buf.toString('latin1', 0, Math.min(buf.length, 200));
  const m = /^<\?xml[^>]*\bencoding\s*=\s*["']([A-Za-z0-9._-]+)["']/.exec(head);
  const declared = m ? m[1].toLowerCase() : 'utf-8';
  if (declared !== 'utf-8' && declared !== 'utf8') return withLabel(buf, declared);
  // Declared (or assumed) UTF-8 that is not: read it with the code page.
  return isUtf8(buf) ? buf.toString('utf8') : decodeMixed(buf);
}

function decodeMixed(buf) {
  const parts = [];
  let start = 0;
  for (let i = 0; i <= buf.length; i++) {
    if (i === buf.length || buf[i] === 0x0a) {
      const line = buf.subarray(start, Math.min(i + 1, buf.length));
      parts.push(isUtf8(line) ? line.toString('utf8') : withLabel(line, FALLBACK));
      start = i + 1;
    }
  }
  return parts.join('');
}

function decodeText(buf) {
  const b = bom(buf);
  if (b) return withLabel(buf.subarray(b.skip), b.encoding);
  return isUtf8(buf) ? buf.toString('utf8') : decodeMixed(buf);
}

module.exports = { decodeXml, decodeText };
