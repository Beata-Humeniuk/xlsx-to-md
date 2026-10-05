'use strict';

// Minimal ZIP reader for OOXML packages: central directory, stored and
// deflated entries. Enough for .xlsx; no ZIP64, no encryption.

const zlib = require('zlib');

const EOCD = 0x06054b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;

function findEndOfCentralDirectory(buf) {
  const min = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === EOCD) return i;
  }
  return -1;
}

// Returns Map<name, () => Buffer>; entries are inflated lazily.
function readZip(buf) {
  if (!Buffer.isBuffer(buf) || buf.length < 22) throw new Error('not a zip archive');
  const eocd = findEndOfCentralDirectory(buf);
  if (eocd < 0) throw new Error('not a zip archive');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const entries = new Map();
  for (let n = 0; n < count; n++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== CENTRAL) throw new Error('corrupt zip central directory');
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const compressedSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString(flags & 0x800 ? 'utf8' : 'latin1', p + 46, p + 46 + nameLen);
    p += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith('/')) continue;
    entries.set(name, () => {
      if (buf.readUInt32LE(localOffset) !== LOCAL) throw new Error('corrupt zip entry: ' + name);
      const start = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28);
      const data = buf.subarray(start, start + compressedSize);
      if (method === 0) return Buffer.from(data);
      if (method === 8) return zlib.inflateRawSync(data);
      throw new Error('unsupported zip compression method ' + method + ' in ' + name);
    });
  }
  return entries;
}

module.exports = { readZip };
