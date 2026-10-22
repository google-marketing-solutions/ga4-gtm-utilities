/**
 * @fileoverview Date formatting helpers emulating `Utilities.formatDate`
 * (Java SimpleDateFormat patterns) and a minimal subset of Google Sheets
 * number formats used for display values.
 */
'use strict';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** @type {!Map<string, !Intl.DateTimeFormat>} */
const formatterCache = new Map();

/**
 * Parses fixed-offset zone ids such as 'GMT', 'UTC', 'GMT+2', 'GMT-05:00',
 * 'UTC+0530', 'Z'. Returns the offset in minutes, or null when `tz` is not a
 * fixed-offset id.
 * @param {string} tz
 * @return {?number}
 */
function parseFixedOffset(tz) {
  if (/^(GMT|UTC|Z|Etc\/(GMT|UTC))$/i.test(tz)) return 0;
  const m = /^(?:GMT|UTC)?([+-])(\d{1,2})(?::?(\d{2}))?$/i.exec(tz);
  if (!m) return null;
  const sign = m[1] === '-' ? -1 : 1;
  return sign * (Number(m[2]) * 60 + Number(m[3] || 0));
}

/**
 * Returns a cached Intl formatter for an IANA zone, or null if unsupported.
 * @param {string} tz
 * @return {?Intl.DateTimeFormat}
 */
function getFormatter(tz) {
  if (formatterCache.has(tz)) return formatterCache.get(tz);
  let fmt = null;
  try {
    fmt = new Intl.DateTimeFormat('en-US', {
      timeZone: tz, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric',
      hour: 'numeric', minute: 'numeric', second: 'numeric', weekday: 'short',
      timeZoneName: 'short',
    });
  } catch (e) {
    fmt = null;
  }
  formatterCache.set(tz, fmt);
  return fmt;
}

/**
 * Breaks a timestamp into calendar fields in the given time zone. Unknown zone
 * ids fall back to GMT, like Java's TimeZone.getTimeZone().
 * @param {number} time Epoch milliseconds.
 * @param {string=} tz Time zone id.
 * @return {{year: number, month: number, day: number, hour: number,
 *     minute: number, second: number, ms: number, weekday: number,
 *     offsetMinutes: number, zoneName: string}}
 */
function getZonedParts(time, tz = 'GMT') {
  const ms = ((time % 1000) + 1000) % 1000;
  const fixed = parseFixedOffset(String(tz || 'GMT'));
  const fmt = fixed === null ? getFormatter(String(tz)) : null;
  if (fixed !== null || !fmt) {
    const offset = fixed === null ? 0 : fixed;
    const d = new Date(time + offset * 60000);
    return {
      year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
      hour: d.getUTCHours(), minute: d.getUTCMinutes(), second: d.getUTCSeconds(), ms,
      weekday: d.getUTCDay(), offsetMinutes: offset,
      zoneName: offset === 0 ? 'GMT' : `GMT${formatOffset(offset, true)}`,
    };
  }
  const parts = {};
  for (const p of fmt.formatToParts(new Date(time))) parts[p.type] = p.value;
  const year = Number(parts.year);
  const month = Number(parts.month);
  const day = Number(parts.day);
  const hour = Number(parts.hour) % 24;
  const minute = Number(parts.minute);
  const second = Number(parts.second);
  const asUtc = Date.UTC(year, month - 1, day, hour, minute, second, ms);
  const offsetMinutes = Math.round((asUtc - time) / 60000);
  return {
    year, month, day, hour, minute, second, ms,
    weekday: DAYS.findIndex((d) => d.startsWith(parts.weekday)),
    offsetMinutes,
    zoneName: parts.timeZoneName || 'GMT',
  };
}

/**
 * Formats an offset in minutes as +HHMM or +HH:MM.
 * @param {number} offsetMinutes
 * @param {boolean} colon
 * @return {string}
 */
function formatOffset(offsetMinutes, colon) {
  const sign = offsetMinutes < 0 ? '-' : '+';
  const abs = Math.abs(offsetMinutes);
  const hh = String(Math.floor(abs / 60)).padStart(2, '0');
  const mm = String(abs % 60).padStart(2, '0');
  return colon ? `${sign}${hh}:${mm}` : `${sign}${hh}${mm}`;
}

/**
 * Formats a date using a Java SimpleDateFormat pattern, like
 * `Utilities.formatDate(date, timeZone, pattern)`.
 *
 * Supported letters: G y Y M L d D E u a H k K h m s S z Z X, quoted literals
 * ('...') and escaped quotes ('').
 *
 * @param {!Date|number} date
 * @param {string} timeZone
 * @param {string} pattern
 * @return {string}
 */
function formatDateJava(date, timeZone, pattern) {
  const time = typeof date === 'number' ? date : date.getTime();
  if (Number.isNaN(time)) throw new Error('Invalid date');
  const p = getZonedParts(time, timeZone);
  let out = '';
  let i = 0;
  while (i < pattern.length) {
    const ch = pattern[i];
    if (ch === '\'') {
      if (pattern[i + 1] === '\'') {
        out += '\'';
        i += 2;
        continue;
      }
      const end = pattern.indexOf('\'', i + 1);
      const stop = end === -1 ? pattern.length : end;
      out += pattern.slice(i + 1, stop).replace(/''/g, '\'');
      i = stop + 1;
      continue;
    }
    if (!/[A-Za-z]/.test(ch)) {
      out += ch;
      i++;
      continue;
    }
    let n = 1;
    while (pattern[i + n] === ch) n++;
    i += n;
    const pad = (v, width = n) => String(v).padStart(width, '0');
    switch (ch) {
      case 'G': out += p.year > 0 ? 'AD' : 'BC'; break;
      case 'y':
      case 'Y': out += n === 2 ? pad(p.year % 100, 2) : pad(p.year); break;
      case 'M':
      case 'L':
        if (n >= 4) out += MONTHS[p.month - 1];
        else if (n === 3) out += MONTHS[p.month - 1].slice(0, 3);
        else out += pad(p.month);
        break;
      case 'd': out += pad(p.day); break;
      case 'D': {
        const start = Date.UTC(p.year, 0, 1);
        const current = Date.UTC(p.year, p.month - 1, p.day);
        out += pad(Math.round((current - start) / 86400000) + 1);
        break;
      }
      case 'E': out += n >= 4 ? DAYS[p.weekday] : DAYS[p.weekday].slice(0, 3); break;
      case 'u': out += pad(p.weekday === 0 ? 7 : p.weekday); break;
      case 'a': out += p.hour < 12 ? 'AM' : 'PM'; break;
      case 'H': out += pad(p.hour); break;
      case 'k': out += pad(p.hour === 0 ? 24 : p.hour); break;
      case 'K': out += pad(p.hour % 12); break;
      case 'h': out += pad(p.hour % 12 === 0 ? 12 : p.hour % 12); break;
      case 'm': out += pad(p.minute); break;
      case 's': out += pad(p.second); break;
      case 'S': out += pad(p.ms); break;
      case 'z': out += p.zoneName; break;
      case 'Z': out += formatOffset(p.offsetMinutes, false); break;
      case 'X':
        if (p.offsetMinutes === 0) out += 'Z';
        else if (n === 1) out += formatOffset(p.offsetMinutes, false).slice(0, 3);
        else out += formatOffset(p.offsetMinutes, n >= 3);
        break;
      default:
        throw new Error(`Illegal pattern character '${ch}'`);
    }
  }
  return out;
}

/**
 * Formats a Date for display in a cell using a (subset of a) Sheets number
 * format, e.g. 'yyyy-mm-dd', 'M/d/yyyy H:mm:ss'. When `format` is empty or the
 * automatic default, 'M/d/yyyy H:mm:ss' is used (or 'M/d/yyyy' at midnight).
 * @param {!Date} date
 * @param {string} timeZone
 * @param {?string} format
 * @return {string}
 */
function formatDateSheets(date, timeZone, format) {
  const p = getZonedParts(date.getTime(), timeZone);
  let fmt = format;
  if (!fmt || fmt === '0.###############' || !/[ymdhs]/i.test(fmt)) {
    fmt = (p.hour || p.minute || p.second) ? 'M/d/yyyy H:mm:ss' : 'M/d/yyyy';
  }
  // Tokenize; 'm' means minutes when it follows an hour token or precedes seconds.
  const tokens = fmt.match(/(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|AM\/PM|"[^"]*"|\\.|.)/gi) || [];
  let out = '';
  tokens.forEach((tok, idx) => {
    const lower = tok.toLowerCase();
    const prev = tokens.slice(0, idx).reverse().find((t) => /^[a-z]/i.test(t));
    const next = tokens.slice(idx + 1).find((t) => /^[a-z]/i.test(t));
    const isMinute = (lower === 'mm' || lower === 'm') &&
        ((prev && /^h/i.test(prev)) || (next && /^s/i.test(next)));
    const pad2 = (v) => String(v).padStart(2, '0');
    switch (lower) {
      case 'yyyy': out += String(p.year); break;
      case 'yy': out += pad2(p.year % 100); break;
      case 'mmmm': out += MONTHS[p.month - 1]; break;
      case 'mmm': out += MONTHS[p.month - 1].slice(0, 3); break;
      case 'mm': out += isMinute ? pad2(p.minute) : pad2(p.month); break;
      case 'm': out += isMinute ? String(p.minute) : String(p.month); break;
      case 'dddd': out += DAYS[p.weekday]; break;
      case 'ddd': out += DAYS[p.weekday].slice(0, 3); break;
      case 'dd': out += pad2(p.day); break;
      case 'd': out += String(p.day); break;
      case 'hh': out += pad2(fmt.match(/am\/pm/i) ? (p.hour % 12 || 12) : p.hour); break;
      case 'h': out += String(fmt.match(/am\/pm/i) ? (p.hour % 12 || 12) : p.hour); break;
      case 'ss': out += pad2(p.second); break;
      case 's': out += String(p.second); break;
      case 'am/pm': out += p.hour < 12 ? 'AM' : 'PM'; break;
      default:
        if (tok.startsWith('"')) out += tok.slice(1, -1);
        else if (tok.startsWith('\\')) out += tok.slice(1);
        else out += tok;
    }
  });
  return out;
}

/**
 * Formats a number for display using a small subset of Sheets number formats:
 * '@', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', optional literal
 * prefix/suffix such as '$'. Anything else uses the automatic format.
 * @param {number} value
 * @param {?string} format
 * @return {string}
 */
function formatNumberSheets(value, format) {
  if (!Number.isFinite(value)) return String(value);
  const auto = () => {
    if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
    const s = String(Number(value.toPrecision(15)));
    return s.replace('e', 'E');
  };
  if (!format || format === '0.###############' || format === 'General') return auto();
  if (format === '@') return String(value);
  const m = /^([^0#.,%]*)([#,0]*0)(?:\.(0+))?(%?)([^0#.,%]*)$/.exec(format);
  if (!m) return auto();
  const [, prefix, intPart, decimals = '', percent, suffix] = m;
  let v = percent ? value * 100 : value;
  const digits = decimals.length;
  let s = Math.abs(v).toFixed(digits);
  if (intPart.includes(',')) {
    const [i, d] = s.split('.');
    s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d !== undefined ? `.${d}` : '');
  }
  const negative = v < 0 && Number(s.replace(/,/g, '')) !== 0;
  return `${negative ? '-' : ''}${prefix.replace(/"/g, '')}${s}${percent}${suffix.replace(/"/g, '')}`;
}

/**
 * Parses a date string with a numeric Java SimpleDateFormat pattern, like
 * `Utilities.parseDate(dateString, timeZone, pattern)`. Supported letters:
 * y M d H m s S and literal characters. Throws if the text does not match.
 * @param {string} text
 * @param {string} timeZone
 * @param {string} pattern
 * @return {number} Milliseconds since the epoch.
 */
function parseDateJava(text, timeZone, pattern) {
  const fields = {y: 1970, M: 1, d: 1, H: 0, m: 0, s: 0, S: 0};
  let regex = '^';
  const order = [];
  const tokens = String(pattern).match(/(y+|M+|d+|H+|m+|s+|S+|'[^']*'|.)/g) || [];
  for (const token of tokens) {
    const letter = token[0];
    if ('yMdHmsS'.indexOf(letter) !== -1) {
      regex += letter === 'y' && token.length >= 4 ? '(\\d{4})' : '(\\d{1,' + Math.max(2, token.length) + '})';
      order.push(letter);
    } else if (letter === "'") {
      regex += token.slice(1, -1).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    } else {
      regex += token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
  }
  const match = new RegExp(regex).exec(String(text));
  if (!match) throw new Error(`Unparseable date: "${text}"`);
  order.forEach((letter, index) => {
    fields[letter] = Number(match[index + 1]);
  });
  const guess = Date.UTC(fields.y, fields.M - 1, fields.d, fields.H, fields.m, fields.s, fields.S);
  let offset = getZonedParts(guess, timeZone).offsetMinutes;
  let time = guess - offset * 60000;
  const corrected = getZonedParts(time, timeZone).offsetMinutes;
  if (corrected !== offset) time = guess - corrected * 60000;
  return time;
}

module.exports = {
  parseDateJava,
  formatDateJava,
  formatDateSheets,
  formatNumberSheets,
  getZonedParts,
  parseFixedOffset,
};
