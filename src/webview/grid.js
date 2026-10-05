'use strict';

// The sheet grid: a virtualized view (only the cells on screen exist in the
// DOM) with Excel-like selection — drag to select, Ctrl/⌘ + drag to add a
// range, Shift + click to extend, column and row headers, double-click for
// the block of data around a cell.

const { columnLetter } = require('../core/address');

const ROW_H = 22;
const HEAD_H = 24;
const ROW_HEAD_W = 52;
const EXTRA_ROWS = 8;
const EXTRA_COLS = 3;

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

class Grid {
  // host: element to fill. callbacks: { onSelect(ranges, activeIndex, done),
  // onRegion(row, col) } — onSelect is called while dragging (done=false)
  // and once more when the mouse is released (done=true).
  constructor(host, callbacks) {
    this.host = host;
    this.cb = callbacks;
    this.sheet = null;
    this.ranges = [];
    this.active = -1;
    this.decor = { headerRows: new Set(), keptRows: null, excludedCols: new Set(), dataRanges: [] };
    this.viewport = el('div', 'grid-viewport');
    this.viewport.tabIndex = 0;
    this.canvas = el('div', 'grid-canvas');
    this.layer = el('div', 'grid-layer');
    this.viewport.appendChild(this.canvas);
    this.viewport.appendChild(this.layer);
    host.appendChild(this.viewport);
    this.viewport.addEventListener('scroll', () => this.schedule());
    this.viewport.addEventListener('mousedown', (e) => this.onDown(e));
    this.viewport.addEventListener('dblclick', (e) => this.onDouble(e));
    window.addEventListener('mousemove', (e) => this.onMove(e));
    window.addEventListener('mouseup', () => this.onUp());
    new ResizeObserver(() => this.schedule()).observe(this.viewport);
    this.drag = null;
    this.frame = 0;
  }

  setSheet(sheet) {
    this.sheet = sheet;
    this.hiddenRows = new Set(sheet.hiddenRows);
    this.hiddenCols = new Set(sheet.hiddenCols);
    this.mergeStart = new Map();
    this.mergeCover = new Map();
    for (const m of sheet.merges) {
      this.mergeStart.set(m.top * 16384 + m.left, m);
      for (let r = m.top; r <= m.bottom; r++) {
        for (let c = m.left; c <= m.right; c++) if (r !== m.top || c !== m.left) this.mergeCover.set(r * 16384 + c, m);
      }
    }
    this.nRows = sheet.rows + EXTRA_ROWS;
    this.nCols = Math.max(sheet.cols + EXTRA_COLS, 6);
    this.widths = [];
    for (let c = 0; c < this.nCols; c++) {
      let chars = 3;
      const limit = Math.min(sheet.rows, 300);
      for (let r = 0; r < limit; r++) {
        const t = c < sheet.cols ? sheet.text[r][c] : '';
        if (t.length <= chars) continue;
        // A title merged across columns does not widen its first column.
        const m = this.mergeStart.get(r * 16384 + c);
        if (m && m.right > m.left) continue;
        chars = t.length;
      }
      this.widths.push(Math.max(56, Math.min(260, Math.round(chars * 7.2 + 16))));
    }
    this.colX = [0];
    for (let c = 0; c < this.nCols; c++) this.colX.push(this.colX[c] + this.widths[c]);
    this.canvas.style.width = ROW_HEAD_W + this.colX[this.nCols] + 'px';
    this.canvas.style.height = HEAD_H + this.nRows * ROW_H + 'px';
    this.viewport.scrollTop = 0;
    this.viewport.scrollLeft = 0;
    this.schedule();
  }

  setSelection(ranges, active) {
    this.ranges = ranges;
    this.active = active;
    this.schedule();
  }

  setDecor(decor) {
    this.decor = decor;
    this.schedule();
  }

  scrollTo(range) {
    if (!this.sheet || !range) return;
    const y = range.top * ROW_H;
    const x = this.colX[Math.min(range.left, this.nCols - 1)];
    const v = this.viewport;
    if (y < v.scrollTop || y > v.scrollTop + v.clientHeight - HEAD_H - ROW_H * 2) v.scrollTop = Math.max(0, y - ROW_H * 2);
    if (x < v.scrollLeft || x > v.scrollLeft + v.clientWidth - ROW_HEAD_W - 80) v.scrollLeft = Math.max(0, x - 40);
  }

  schedule() {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.render();
    });
  }

  colAt(x) {
    let lo = 0;
    let hi = this.nCols - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.colX[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  // Cell under a mouse event: { row, col, area: 'cell' | 'colhead' | 'rowhead' | 'corner' }.
  hit(e) {
    const rect = this.viewport.getBoundingClientRect();
    const vx = e.clientX - rect.left;
    const vy = e.clientY - rect.top;
    const x = vx + this.viewport.scrollLeft - ROW_HEAD_W;
    const y = vy + this.viewport.scrollTop - HEAD_H;
    const col = Math.max(0, Math.min(this.nCols - 1, this.colAt(Math.max(0, x))));
    const row = Math.max(0, Math.min(this.nRows - 1, Math.floor(Math.max(0, y) / ROW_H)));
    let area = 'cell';
    if (vy < HEAD_H && vx < ROW_HEAD_W) area = 'corner';
    else if (vy < HEAD_H) area = 'colhead';
    else if (vx < ROW_HEAD_W) area = 'rowhead';
    return { row, col, area, vx, vy, rect };
  }

  lastRow() {
    return Math.max(0, this.sheet.rows - 1);
  }

  lastCol() {
    return Math.max(0, this.sheet.cols - 1);
  }

  rangeFor(anchor, h, area) {
    if (area === 'colhead') {
      return { top: 0, left: Math.min(anchor.col, h.col), bottom: this.lastRow(), right: Math.max(anchor.col, h.col) };
    }
    if (area === 'rowhead') {
      return { top: Math.min(anchor.row, h.row), left: 0, bottom: Math.max(anchor.row, h.row), right: this.lastCol() };
    }
    return {
      top: Math.min(anchor.row, h.row),
      left: Math.min(anchor.col, h.col),
      bottom: Math.max(anchor.row, h.row),
      right: Math.max(anchor.col, h.col),
    };
  }

  onDown(e) {
    if (!this.sheet || e.button !== 0) return;
    const h = this.hit(e);
    if (h.vx > this.viewport.clientWidth || h.vy > this.viewport.clientHeight) return;
    e.preventDefault();
    this.viewport.focus();
    if (h.area === 'corner') {
      const all = { top: 0, left: 0, bottom: this.lastRow(), right: this.lastCol() };
      this.ranges = [all];
      this.active = 0;
      this.schedule();
      this.cb.onSelect(this.ranges, this.active, true);
      return;
    }
    const additive = e.ctrlKey || e.metaKey;
    let anchor = { row: h.row, col: h.col };
    if (e.shiftKey && this.active >= 0 && this.ranges[this.active]) {
      const r = this.ranges[this.active];
      anchor = r.anchor || { row: r.top, col: r.left };
    }
    const range = this.rangeFor(anchor, h, h.area);
    range.anchor = anchor;
    if (additive) {
      this.ranges = this.ranges.concat([range]);
      this.active = this.ranges.length - 1;
    } else if (e.shiftKey && this.active >= 0) {
      this.ranges = this.ranges.slice();
      this.ranges[this.active] = Object.assign({}, this.ranges[this.active], range);
    } else {
      this.ranges = [range];
      this.active = 0;
    }
    this.drag = { anchor, area: h.area, lastEvent: e };
    this.schedule();
    this.cb.onSelect(this.ranges, this.active, false);
    this.startAutoScroll();
  }

  onMove(e) {
    if (!this.drag) return;
    this.drag.lastEvent = e;
    this.extendTo(e);
  }

  extendTo(e) {
    const h = this.hit(e);
    const range = this.rangeFor(this.drag.anchor, h, this.drag.area);
    const cur = this.ranges[this.active];
    if (cur && cur.top === range.top && cur.left === range.left && cur.bottom === range.bottom && cur.right === range.right) return;
    this.ranges = this.ranges.slice();
    this.ranges[this.active] = Object.assign({}, cur, range);
    this.schedule();
    this.cb.onSelect(this.ranges, this.active, false);
  }

  startAutoScroll() {
    clearInterval(this.autoTimer);
    this.autoTimer = setInterval(() => {
      if (!this.drag) {
        clearInterval(this.autoTimer);
        return;
      }
      const e = this.drag.lastEvent;
      const rect = this.viewport.getBoundingClientRect();
      let dx = 0;
      let dy = 0;
      if (e.clientY > rect.bottom - 8) dy = ROW_H * 2;
      else if (e.clientY < rect.top + HEAD_H && this.drag.area !== 'colhead') dy = -ROW_H * 2;
      if (e.clientX > rect.right - 8) dx = 60;
      else if (e.clientX < rect.left + ROW_HEAD_W && this.drag.area !== 'rowhead') dx = -60;
      if (dx || dy) {
        this.viewport.scrollLeft += dx;
        this.viewport.scrollTop += dy;
        this.extendTo(e);
      }
    }, 50);
  }

  onUp() {
    if (!this.drag) return;
    this.drag = null;
    clearInterval(this.autoTimer);
    this.cb.onSelect(this.ranges, this.active, true);
  }

  onDouble(e) {
    if (!this.sheet) return;
    const h = this.hit(e);
    if (h.area !== 'cell') return;
    this.cb.onRegion(h.row, h.col, e.ctrlKey || e.metaKey);
  }

  cellClass(r, c) {
    let cls = 'cell';
    if (this.hiddenRows.has(r) || this.hiddenCols.has(c)) cls += ' is-hidden';
    let inRange = false;
    for (const rg of this.ranges) {
      if (r >= rg.top && r <= rg.bottom && c >= rg.left && c <= rg.right) {
        inRange = true;
        break;
      }
    }
    if (inRange) {
      cls += ' in-sel';
      const d = this.decor;
      if (d.headerRows.has(r) && d.dataRanges.some((rg) => rg.top === r && c >= rg.left && c <= rg.right)) cls += ' is-head';
      else if (d.keptRows && !d.keptRows.has(r)) cls += ' is-out';
      if (d.excludedCols.has(c)) cls += ' col-out';
    }
    if (r < this.sheet.rows && c < this.sheet.cols) {
      const k = this.sheet.kinds[r][c];
      if (k === 'n' || k === 'd') cls += ' num';
      else if (k === 'b' || k === 'e') cls += ' center';
    }
    return cls;
  }

  render() {
    if (!this.sheet) return;
    const v = this.viewport;
    const sheet = this.sheet;
    const top = v.scrollTop;
    const left = v.scrollLeft;
    const height = v.clientHeight;
    const width = v.clientWidth;
    const r0 = Math.max(0, Math.floor(top / ROW_H) - 1);
    const r1 = Math.min(this.nRows - 1, Math.ceil((top + height) / ROW_H) + 1);
    const c0 = Math.max(0, this.colAt(left) - 1);
    const c1 = Math.min(this.nCols - 1, this.colAt(left + width) + 1);

    const frag = document.createDocumentFragment();
    const done = new Set();
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        let rr = r;
        let cc = c;
        let span = null;
        const cover = this.mergeCover.get(r * 16384 + c);
        if (cover) {
          rr = cover.top;
          cc = cover.left;
          span = cover;
        } else if (this.mergeStart.has(r * 16384 + c)) {
          span = this.mergeStart.get(r * 16384 + c);
        }
        const key = rr * 16384 + cc;
        if (done.has(key)) continue;
        done.add(key);
        const text = rr < sheet.rows && cc < sheet.cols ? sheet.text[rr][cc] : '';
        const d = el('div', this.cellClass(rr, cc) + (span ? ' merged' : ''), text);
        const x = this.colX[cc];
        const y = rr * ROW_H;
        const w = span ? this.colX[Math.min(span.right + 1, this.nCols)] - x : this.widths[cc];
        const hgt = span ? (Math.min(span.bottom, this.nRows - 1) - rr + 1) * ROW_H : ROW_H;
        d.style.transform = 'translate(' + (ROW_HEAD_W + x) + 'px,' + (HEAD_H + y) + 'px)';
        d.style.width = w + 'px';
        d.style.height = hgt + 'px';
        if (text.length > 20) d.title = text;
        frag.appendChild(d);
      }
    }

    // Selection outlines.
    this.ranges.forEach((rg, i) => {
      if (rg.bottom < r0 - 2 && rg.top < r0) return;
      const o = el('div', 'sel-outline' + (i === this.active ? ' is-active' : ''));
      const x = this.colX[Math.min(rg.left, this.nCols)];
      const x2 = this.colX[Math.min(rg.right + 1, this.nCols)];
      o.style.transform = 'translate(' + (ROW_HEAD_W + x) + 'px,' + (HEAD_H + rg.top * ROW_H) + 'px)';
      o.style.width = x2 - x + 'px';
      o.style.height = (Math.min(rg.bottom, this.nRows - 1) - rg.top + 1) * ROW_H + 'px';
      frag.appendChild(o);
    });

    // Column headers (stick to the top), row headers (stick to the left).
    const selCols = new Set();
    const selRows = new Set();
    for (const rg of this.ranges) {
      for (let c = Math.max(rg.left, c0); c <= Math.min(rg.right, c1); c++) selCols.add(c);
      for (let r = Math.max(rg.top, r0); r <= Math.min(rg.bottom, r1); r++) selRows.add(r);
    }
    for (let c = c0; c <= c1; c++) {
      const h = el('div', 'col-head' + (selCols.has(c) ? ' is-sel' : '') + (this.hiddenCols.has(c) ? ' is-hidden' : ''), columnLetter(c));
      h.style.transform = 'translate(' + (ROW_HEAD_W + this.colX[c]) + 'px,' + top + 'px)';
      h.style.width = this.widths[c] + 'px';
      frag.appendChild(h);
    }
    for (let r = r0; r <= r1; r++) {
      const h = el('div', 'row-head' + (selRows.has(r) ? ' is-sel' : '') + (this.hiddenRows.has(r) ? ' is-hidden' : ''), String(r + 1));
      h.style.transform = 'translate(' + left + 'px,' + (HEAD_H + r * ROW_H) + 'px)';
      frag.appendChild(h);
    }
    const corner = el('div', 'corner');
    corner.style.transform = 'translate(' + left + 'px,' + top + 'px)';
    frag.appendChild(corner);

    this.layer.replaceChildren(frag);
  }
}

module.exports = { Grid };
