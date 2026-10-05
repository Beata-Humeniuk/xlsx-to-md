'use strict';

// Excel number formats: the text a numeric cell shows in Excel, built from its
// value and format code. Covers sections (positive;negative;zero), digit
// placeholders 0 # ?, grouping and scaling commas, percent, scientific
// notation, fractions, literals, currency tags and dates/times. Locale-bound
// built-in formats (short date, currency) are written in a neutral form: ISO
// dates and plain numbers.

const BUILTIN = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '#,##0;-#,##0',
  6: '#,##0;-#,##0',
  7: '#,##0.00;-#,##0.00',
  8: '#,##0.00;-#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'yyyy-mm-dd',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'yyyy-mm-dd hh:mm',
  27: 'yyyy-mm-dd',
  28: 'yyyy-mm-dd',
  29: 'yyyy-mm-dd',
  30: 'yyyy-mm-dd',
  31: 'yyyy-mm-dd',
  32: 'h:mm:ss',
  33: 'h:mm:ss',
  34: 'h:mm:ss',
  35: 'h:mm:ss',
  36: 'yyyy-mm-dd',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;(#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;(#,##0.00)',
  41: '#,##0',
  42: '#,##0',
  43: '#,##0.00',
  44: '#,##0.00',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
  50: 'yyyy-mm-dd',
  51: 'yyyy-mm-dd',
  52: 'yyyy-mm-dd',
  53: 'yyyy-mm-dd',
  54: 'yyyy-mm-dd',
  55: 'yyyy-mm-dd',
  56: 'yyyy-mm-dd',
  57: 'yyyy-mm-dd',
  58: 'yyyy-mm-dd',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function builtinFormat(id) {
  return BUILTIN[id] !== undefined ? BUILTIN[id] : null;
}

// --- tokenizer --------------------------------------------------------------

function splitSections(code) {
  const out = [];
  let cur = '';
  let quote = false;
  let bracket = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (quote) {
      cur += ch;
      if (ch === '"') quote = false;
    } else if (ch === '\\' && i + 1 < code.length) {
      cur += ch + code[++i];
    } else if (ch === '"') {
      quote = true;
      cur += ch;
    } else if (ch === '[') {
      bracket = true;
      cur += ch;
    } else if (ch === ']') {
      bracket = false;
      cur += ch;
    } else if (ch === ';' && !bracket) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

const DATE_RE = /^(yyyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|e|bb|b)/i;

function tokenize(section) {
  const tokens = [];
  let condition = null;
  let i = 0;
  const lit = (text) => {
    const last = tokens[tokens.length - 1];
    if (last && last.type === 'lit') last.text += text;
    else tokens.push({ type: 'lit', text });
  };
  while (i < section.length) {
    const ch = section[i];
    const rest = section.slice(i);
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      lit(section.slice(i + 1, end < 0 ? section.length : end));
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '\\') {
      lit(section[i + 1] || '');
      i += 2;
    } else if (ch === '_') {
      lit(' ');
      i += 2;
    } else if (ch === '*') {
      i += 2;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      const inner = section.slice(i + 1, end < 0 ? section.length : end);
      i = end < 0 ? section.length : end + 1;
      const elapsed = /^(h+|m+|s+)$/i.exec(inner);
      const cond = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner);
      if (elapsed) tokens.push({ type: 'date', code: inner.toLowerCase(), elapsed: true });
      else if (cond) condition = { op: cond[1], value: parseFloat(cond[2]) };
      else if (inner[0] === '$') {
        const symbol = inner.slice(1).split('-')[0];
        if (symbol) lit(symbol);
      }
      // Colors, locale and calendar tags change nothing in plain text.
    } else if (/^general/i.test(rest)) {
      tokens.push({ type: 'general' });
      i += 7;
    } else if (/^(am\/pm|a\/p)/i.test(rest)) {
      const m = /^(am\/pm|a\/p)/i.exec(rest)[0];
      tokens.push({ type: 'ampm', code: m });
      i += m.length;
    } else if (/^e[+-]/i.test(rest)) {
      tokens.push({ type: 'exp', sign: rest[1] });
      i += 2;
    } else if (DATE_RE.test(rest) && !(ch.toLowerCase() === 'e' && !/[ymdhs]/i.test(section))) {
      const m = DATE_RE.exec(rest)[0];
      tokens.push({ type: 'date', code: m.toLowerCase() });
      i += m.length;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      tokens.push({ type: 'ph', ch });
      i++;
    } else if (ch === '.') {
      tokens.push({ type: 'dot' });
      i++;
    } else if (ch === ',') {
      tokens.push({ type: 'comma' });
      i++;
    } else if (ch === '%') {
      tokens.push({ type: 'pct' });
      i++;
    } else if (ch === '/') {
      tokens.push({ type: 'slash' });
      i++;
    } else if (ch === '@') {
      tokens.push({ type: 'text' });
      i++;
    } else if (/[1-9]/.test(ch) && tokens.some((t) => t.type === 'slash')) {
      // A fixed denominator such as "# ?/8".
      const m = /^\d+/.exec(rest)[0];
      tokens.push({ type: 'denom', value: parseInt(m, 10) });
      i += m.length;
    } else {
      lit(ch);
      i++;
    }
  }
  return { tokens, condition };
}

// --- general ----------------------------------------------------------------

function formatGeneral(value) {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e21) return String(value);
  const n = parseFloat(value.toPrecision(15));
  return String(n);
}

// --- dates ------------------------------------------------------------------

// Excel serial date → UTC parts. The 1900 system counts the non-existent
// 1900-02-29 (serial 60); dates before it are one day off from a plain epoch.
function serialToParts(serial, date1904) {
  let days = Math.floor(serial);
  let seconds = Math.round((serial - days) * 86400 * 1000) / 1000;
  if (seconds >= 86400) {
    days += 1;
    seconds -= 86400;
  }
  let epoch;
  if (date1904) epoch = Date.UTC(1904, 0, 1);
  else if (days >= 61) epoch = Date.UTC(1899, 11, 30);
  else if (days === 60) return partsOf(Date.UTC(1900, 1, 28), seconds, true);
  else epoch = Date.UTC(1899, 11, 31);
  return partsOf(epoch + days * 86400000, seconds, false);
}

function partsOf(ms, seconds, leapBug) {
  const d = new Date(ms);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: leapBug ? 29 : d.getUTCDate(),
    weekday: d.getUTCDay(),
    seconds,
  };
}

function pad(n, width) {
  const s = String(n);
  return s.length >= width ? s : '0'.repeat(width - s.length) + s;
}

function formatDate(value, tokens, date1904) {
  const hasAmPm = tokens.some((t) => t.type === 'ampm');
  const subDigits = (() => {
    // "ss.0" / "ss.00": fractions of a second.
    for (let i = 0; i < tokens.length - 1; i++) {
      if (tokens[i].type === 'date' && tokens[i].code[0] === 's' && tokens[i + 1].type === 'dot') {
        let n = 0;
        for (let j = i + 2; j < tokens.length && tokens[j].type === 'ph' && tokens[j].ch === '0'; j++) n++;
        return n;
      }
    }
    return 0;
  })();
  const p = serialToParts(value, date1904);
  const totalSeconds = subDigits ? p.seconds : Math.round(p.seconds);
  const whole = Math.floor(totalSeconds);
  const hour = Math.floor(whole / 3600) % 24;
  const minute = Math.floor((whole % 3600) / 60);
  const second = whole % 60;
  const frac = totalSeconds - whole;

  const dateTokens = tokens.map((t, i) => ({ t, i })).filter((x) => x.t.type === 'date');
  const isMinute = (k) => {
    const prev = dateTokens[k - 1];
    const next = dateTokens[k + 1];
    return (prev && prev.t.code[0] === 'h') || (next && next.t.code[0] === 's');
  };
  const minuteAt = new Set();
  dateTokens.forEach((x, k) => {
    if (x.t.code[0] === 'm' && !x.t.elapsed && x.t.code.length <= 2 && isMinute(k)) minuteAt.add(x.i);
  });

  let out = '';
  let skipFraction = 0;
  tokens.forEach((t, i) => {
    if (skipFraction > 0) {
      skipFraction--;
      return;
    }
    if (t.type === 'lit') out += t.text;
    else if (t.type === 'ampm') {
      const pm = hour >= 12;
      out += t.code.length > 3 ? (pm ? 'PM' : 'AM') : (pm ? 'P' : 'A');
    } else if (t.type === 'date') {
      const c = t.code;
      if (t.elapsed) {
        const total = Math.round(value * 86400);
        if (c[0] === 'h') out += pad(Math.floor(total / 3600), c.length);
        else if (c[0] === 'm') out += pad(Math.floor(total / 60), c.length);
        else out += pad(total, c.length);
      } else if (c[0] === 'y' || c[0] === 'e') {
        out += c.length <= 2 && c[0] === 'y' ? pad(p.year % 100, 2) : String(p.year);
      } else if (c[0] === 'b') {
        out += String(p.year + 543);
      } else if (c[0] === 'm' && minuteAt.has(i)) {
        out += c.length === 2 ? pad(minute, 2) : String(minute);
      } else if (c[0] === 'm') {
        if (c.length === 1) out += String(p.month);
        else if (c.length === 2) out += pad(p.month, 2);
        else if (c.length === 3) out += MONTHS[p.month - 1].slice(0, 3);
        else if (c.length === 4) out += MONTHS[p.month - 1];
        else out += MONTHS[p.month - 1][0];
      } else if (c[0] === 'd') {
        if (c.length === 1) out += String(p.day);
        else if (c.length === 2) out += pad(p.day, 2);
        else if (c.length === 3) out += DAYS[p.weekday].slice(0, 3);
        else out += DAYS[p.weekday];
      } else if (c[0] === 'h') {
        let h = hour;
        if (hasAmPm) h = h % 12 === 0 ? 12 : h % 12;
        out += c.length === 2 ? pad(h, 2) : String(h);
      } else if (c[0] === 's') {
        out += c.length === 2 ? pad(second, 2) : String(second);
        if (subDigits && tokens[i + 1] && tokens[i + 1].type === 'dot') {
          out += '.' + pad(Math.round(frac * Math.pow(10, subDigits)), subDigits).slice(0, subDigits);
          skipFraction = 1 + subDigits;
        }
      }
    } else if (t.type === 'dot') out += '.';
    else if (t.type === 'comma') out += ',';
    else if (t.type === 'pct') out += '%';
    else if (t.type === 'slash') out += '/';
    else if (t.type === 'ph') out += t.ch === '?' ? ' ' : t.ch === '0' ? '0' : '';
    else if (t.type === 'general') out += formatGeneral(value);
  });
  return out;
}

function isDateSection(tokens) {
  return tokens.some((t) => t.type === 'date' || t.type === 'ampm');
}

// --- numbers ----------------------------------------------------------------

function roundTo(value, decimals) {
  if (decimals > 15) decimals = 15;
  const shifted = Math.round(Number(value + 'e' + decimals));
  return Number(shifted + 'e-' + decimals);
}

function groupThousands(digits) {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// Writes `digits` into the integer placeholders from the right, the way
// Excel does: placeholders take one digit each, the leftmost one takes all
// remaining digits, literals between placeholders stay where they are.
function fillInteger(tokens, digits) {
  const phCount = tokens.filter((t) => t.type === 'ph').length;
  if (!phCount) return tokens.map((t) => (t.type === 'lit' ? t.text : '')).join('') + digits;
  let rest = digits;
  let seen = 0;
  const parts = new Array(tokens.length).fill('');
  for (let i = tokens.length - 1; i >= 0; i--) {
    const t = tokens[i];
    if (t.type === 'lit') {
      parts[i] = t.text;
      continue;
    }
    if (t.type !== 'ph') continue;
    seen++;
    if (seen === phCount) {
      if (rest.length) parts[i] = rest;
      else parts[i] = t.ch === '0' ? '0' : t.ch === '?' ? ' ' : '';
      rest = '';
    } else if (rest.length) {
      parts[i] = rest[rest.length - 1];
      rest = rest.slice(0, -1);
    } else {
      parts[i] = t.ch === '0' ? '0' : t.ch === '?' ? ' ' : '';
    }
  }
  return parts.join('');
}

function fillFraction(tokens, digits) {
  let k = 0;
  let out = '';
  for (const t of tokens) {
    if (t.type === 'lit') out += t.text;
    else if (t.type === 'ph') out += digits[k++] || '';
  }
  return out;
}

function fractionDigits(phs, digits) {
  // Trailing zeros: '#' drops them, '?' turns them into spaces, '0' keeps them.
  const chars = digits.split('');
  for (let i = phs.length - 1; i >= 0; i--) {
    if (chars[i] !== '0') break;
    if (phs[i].ch === '#') chars[i] = '';
    else if (phs[i].ch === '?') chars[i] = ' ';
    else break;
  }
  return chars;
}

function bestFraction(x, maxDen) {
  let best = { n: Math.round(x), d: 1 };
  let bestErr = Math.abs(x - best.n);
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(x * d);
    const err = Math.abs(x - n / d);
    if (err < bestErr - 1e-12) {
      best = { n, d };
      bestErr = err;
    }
  }
  return best;
}

function formatFraction(value, tokens) {
  const slash = tokens.findIndex((t) => t.type === 'slash');
  const before = tokens.slice(0, slash);
  const after = tokens.slice(slash + 1);
  const fixed = after.find((t) => t.type === 'denom');
  const denPh = after.filter((t) => t.type === 'ph').length || 1;
  // A space-separated group before the numerator is the whole-number part.
  let wholeEnd = -1;
  for (let i = before.length - 1; i >= 0; i--) {
    if (before[i].type === 'lit' && /\s/.test(before[i].text)) {
      wholeEnd = i;
      break;
    }
  }
  const wholeTokens = wholeEnd >= 0 ? before.slice(0, wholeEnd) : [];
  const hasWhole = wholeTokens.some((t) => t.type === 'ph');
  const whole = hasWhole ? Math.floor(value) : 0;
  const rest = value - whole;
  const frac = fixed
    ? { n: Math.round(rest * fixed.value), d: fixed.value }
    : bestFraction(rest, Math.pow(10, denPh) - 1);
  let w = whole;
  if (frac.n === frac.d && hasWhole) {
    w += 1;
    frac.n = 0;
  }
  let out = '';
  if (hasWhole && (w !== 0 || frac.n === 0)) out += String(w);
  if (frac.n !== 0 || !hasWhole) {
    if (out) out += ' ';
    out += frac.n + '/' + frac.d;
  }
  return out || '0';
}

function formatNumberSection(value, tokens) {
  if (tokens.some((t) => t.type === 'general')) {
    return tokens.map((t) => (t.type === 'lit' ? t.text : t.type === 'general' ? formatGeneral(value) : '')).join('');
  }
  if (tokens.some((t) => t.type === 'slash')) {
    const lits = (list) => list.filter((t) => t.type === 'lit').map((t) => t.text).join('');
    const first = tokens.findIndex((t) => t.type === 'ph');
    const prefix = first > 0 ? lits(tokens.slice(0, first)) : '';
    return prefix.trim() + formatFraction(value, tokens);
  }
  if (!tokens.some((t) => t.type === 'ph')) {
    return tokens.map((t) => (t.type === 'lit' ? t.text : t.type === 'pct' ? '%' : '')).join('');
  }

  let v = value;
  for (const t of tokens) if (t.type === 'pct') v *= 100;

  const expAt = tokens.findIndex((t) => t.type === 'exp');
  const mantissa = expAt >= 0 ? tokens.slice(0, expAt) : tokens;
  const expTokens = expAt >= 0 ? tokens.slice(expAt + 1) : [];
  const dotAt = mantissa.findIndex((t) => t.type === 'dot');
  let intTokens = dotAt >= 0 ? mantissa.slice(0, dotAt) : mantissa;
  const fracTokens = dotAt >= 0 ? mantissa.slice(dotAt + 1) : [];

  // Trailing commas after the last integer placeholder scale by 1000 each.
  const lastPh = intTokens.map((t) => t.type).lastIndexOf('ph');
  let grouping = false;
  const kept = [];
  intTokens.forEach((t, i) => {
    if (t.type === 'comma') {
      if (i > lastPh) v /= 1000;
      else grouping = true;
    } else kept.push(t);
  });
  intTokens = kept;
  const fracPh = fracTokens.filter((t) => t.type === 'ph');
  const decimals = fracPh.length;

  let exponentText = '';
  if (expAt >= 0) {
    const intPh = intTokens.filter((t) => t.type === 'ph').length || 1;
    let exp = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    if (intPh > 1 && intTokens.some((t) => t.type === 'ph' && t.ch === '#')) {
      exp = Math.floor(exp / intPh) * intPh;
    } else {
      exp -= intPh - 1;
    }
    v /= Math.pow(10, exp);
    if (roundTo(Math.abs(v), decimals) >= Math.pow(10, intPh) && intPh === 1) {
      v /= 10;
      exp += 1;
    }
    const minDigits = expTokens.filter((t) => t.type === 'ph' && t.ch === '0').length || 1;
    const sign = exp < 0 ? '-' : expTokens.length && tokens[expAt].sign === '+' ? '+' : '';
    exponentText = 'E' + sign + pad(Math.abs(exp), minDigits);
  }

  const rounded = roundTo(Math.abs(v), decimals);
  const fixed = rounded.toFixed(decimals);
  let [intDigits, fracDigits = ''] = fixed.split('.');
  if (intDigits === '0' && !intTokens.some((t) => t.type === 'ph' && t.ch === '0')) intDigits = '';
  if (grouping) intDigits = groupThousands(intDigits);

  let out = fillInteger(intTokens, intDigits);
  if (dotAt >= 0) {
    out += '.' + fillFraction(fracTokens, fractionDigits(fracPh, fracDigits).join(''));
  }
  out += exponentText;
  // Percent signs go last, where formats put them in practice.
  out += tokens.filter((t) => t.type === 'pct').map(() => '%').join('');
  return out;
}

function conditionHolds(c, value) {
  switch (c.op) {
    case '<': return value < c.value;
    case '<=': return value <= c.value;
    case '>': return value > c.value;
    case '>=': return value >= c.value;
    case '=': return value === c.value;
    case '<>': return value !== c.value;
    default: return false;
  }
}

// value: number; code: format code string (or null for General).
function formatNumber(value, code, date1904) {
  if (!Number.isFinite(value)) return String(value);
  if (!code || /^general$/i.test(code.trim())) return formatGeneral(value);
  const sections = splitSections(code).map(tokenize);
  let section;
  let v = value;
  let explicitSign = false;
  if (sections.some((s) => s.condition)) {
    section = sections.find((s) => s.condition && conditionHolds(s.condition, value))
      || sections.find((s) => !s.condition && !s.tokens.some((t) => t.type === 'text'))
      || sections[0];
    if (value < 0 && section !== sections[0] && section.condition && section.condition.value <= 0) {
      v = Math.abs(value);
      explicitSign = true;
    }
  } else if (sections.length === 1 || (sections.length > 1 && value > 0)) {
    section = sections[0];
  } else if (value < 0) {
    section = sections[1] && sections[1].tokens.length ? sections[1] : sections[0];
    if (section !== sections[0]) {
      v = Math.abs(value);
      explicitSign = true;
    }
  } else {
    section = sections.length >= 3 ? sections[2] : sections[0];
  }

  const tokens = section.tokens;
  if (isDateSection(tokens)) {
    if (value < 0) return formatGeneral(value);
    return formatDate(value, tokens, date1904);
  }
  if (tokens.some((t) => t.type === 'text') && !tokens.some((t) => t.type === 'ph')) {
    return tokens.map((t) => (t.type === 'lit' ? t.text : t.type === 'text' ? formatGeneral(value) : '')).join('');
  }
  let text = formatNumberSection(Math.abs(v), tokens);
  if (!explicitSign && v < 0 && /[1-9]/.test(text)) text = '-' + text;
  return text;
}

// True when the format shows the number as a date or time.
function isDateFormat(code) {
  if (!code) return false;
  const first = splitSections(code)[0];
  return isDateSection(tokenize(first).tokens);
}

module.exports = { formatNumber, formatGeneral, isDateFormat, builtinFormat, serialToParts };
