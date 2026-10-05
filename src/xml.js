'use strict';

// Small XML parser for OOXML parts. Produces { name, attrs, children } nodes
// with text as plain strings. Element and attribute prefixes are rewritten to
// fixed prefixes (s: for SpreadsheetML, r:, rel:) by namespace URI, so the
// reader does not depend on which prefixes a producer happened to choose.

const { decodeXml } = require('./text');

const CANONICAL = {
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main': 's',
  'http://purl.oclc.org/ooxml/spreadsheetml/main': 's',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://www.w3.org/XML/1998/namespace': 'xml',
};

const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

function decode(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (whole, e) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return e in ENTITIES ? ENTITIES[e] : whole;
  });
}

function canonicalName(qname, scope, isAttr) {
  const colon = qname.indexOf(':');
  if (colon < 0) {
    if (isAttr) return qname;
    const uri = scope[''];
    const prefix = uri && CANONICAL[uri];
    return prefix ? prefix + ':' + qname : qname;
  }
  const prefix = qname.slice(0, colon);
  if (prefix === 'xmlns') return qname;
  const uri = scope[prefix];
  const canon = uri ? CANONICAL[uri] : null;
  return (canon || prefix) + ':' + qname.slice(colon + 1);
}

const ATTR = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseXml(text) {
  if (Buffer.isBuffer(text)) text = decodeXml(text);
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const root = { name: '#document', attrs: {}, children: [] };
  const stack = [{ node: root, scope: {} }];
  let i = 0;
  const len = text.length;
  while (i < len) {
    const lt = text.indexOf('<', i);
    const top = stack[stack.length - 1];
    if (lt < 0) {
      if (i < len) top.node.children.push(decode(text.slice(i)));
      break;
    }
    if (lt > i) top.node.children.push(decode(text.slice(i, lt)));
    if (text.startsWith('<!--', lt)) {
      const end = text.indexOf('-->', lt + 4);
      i = end < 0 ? len : end + 3;
    } else if (text.startsWith('<![CDATA[', lt)) {
      const end = text.indexOf(']]>', lt + 9);
      top.node.children.push(text.slice(lt + 9, end < 0 ? len : end));
      i = end < 0 ? len : end + 3;
    } else if (text[lt + 1] === '?' || text[lt + 1] === '!') {
      const end = text.indexOf('>', lt);
      i = end < 0 ? len : end + 1;
    } else if (text[lt + 1] === '/') {
      const end = text.indexOf('>', lt);
      if (stack.length > 1) stack.pop();
      i = end < 0 ? len : end + 1;
    } else {
      // Find the end of the tag, skipping '>' inside quoted attribute values.
      let j = lt + 1;
      let quote = null;
      while (j < len) {
        const ch = text[j];
        if (quote) { if (ch === quote) quote = null; } else if (ch === '"' || ch === "'") quote = ch; else if (ch === '>') break;
        j++;
      }
      let body = text.slice(lt + 1, j);
      const selfClosing = body.endsWith('/');
      if (selfClosing) body = body.slice(0, -1);
      const m = /^[^\s/>]+/.exec(body);
      const qname = m ? m[0] : '';
      const rawAttrs = [];
      const scope = Object.create(top.scope);
      ATTR.lastIndex = qname.length;
      let a;
      while ((a = ATTR.exec(body))) {
        const value = decode(a[3] !== undefined ? a[3] : a[4]);
        if (a[1] === 'xmlns') scope[''] = value;
        else if (a[1].startsWith('xmlns:')) scope[a[1].slice(6)] = value;
        rawAttrs.push([a[1], value]);
      }
      const attrs = {};
      for (const [k, v] of rawAttrs) attrs[canonicalName(k, scope, true)] = v;
      const node = { name: canonicalName(qname, scope, false), attrs, children: [] };
      top.node.children.push(node);
      if (!selfClosing) stack.push({ node, scope });
      i = j + 1;
    }
  }
  return root.children.find((c) => typeof c !== 'string') || null;
}

// --- tree helpers ---------------------------------------------------------

function elements(node) {
  return node ? node.children.filter((c) => typeof c !== 'string') : [];
}

function child(node, name) {
  if (!node) return null;
  for (const c of node.children) if (typeof c !== 'string' && c.name === name) return c;
  return null;
}

function childrenNamed(node, name) {
  return elements(node).filter((c) => c.name === name);
}

function attr(node, name) {
  return node && node.attrs[name] !== undefined ? node.attrs[name] : null;
}

function textOf(node) {
  if (!node) return '';
  let out = '';
  for (const c of node.children) out += typeof c === 'string' ? c : textOf(c);
  return out;
}

function descendants(node, name, out = []) {
  for (const c of elements(node)) {
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
}

module.exports = { parseXml, elements, child, childrenNamed, attr, textOf, descendants };
