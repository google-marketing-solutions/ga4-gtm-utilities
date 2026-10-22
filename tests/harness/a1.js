/**
 * @fileoverview A1-notation helpers shared by the spreadsheet mock and the
 * fixture helpers.
 */
'use strict';

/**
 * Converts a 1-based column index to letters (1 -> A, 27 -> AA).
 * @param {number} col
 * @return {string}
 */
function columnToLetters(col) {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Converts column letters to a 1-based index (A -> 1, AA -> 27).
 * @param {string} letters
 * @return {number}
 */
function lettersToColumn(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/**
 * Formats a cell reference.
 * @param {number} row
 * @param {number} col
 * @return {string}
 */
function cellA1(row, col) {
  return `${columnToLetters(col)}${row}`;
}

/**
 * Formats a rectangular range reference ('A1' for single cells).
 * @param {number} row
 * @param {number} col
 * @param {number=} numRows
 * @param {number=} numCols
 * @return {string}
 */
function rangeA1(row, col, numRows = 1, numCols = 1) {
  const start = cellA1(row, col);
  if (numRows === 1 && numCols === 1) return start;
  return `${start}:${cellA1(row + numRows - 1, col + numCols - 1)}`;
}

/**
 * Parses an A1 reference, optionally sheet-qualified ("'My Sheet'!A1:B2").
 * Open-ended references ('A:C', '2:5', 'A2:C') are resolved against
 * `maxRows`/`maxCols`.
 *
 * @param {string} a1
 * @param {number} maxRows
 * @param {number} maxCols
 * @return {?{sheetName: ?string, row: number, col: number, numRows: number,
 *     numCols: number, hint: ?string}} null when the notation is invalid.
 */
function parseA1(a1, maxRows, maxCols) {
  const m = /^(?:(?:'((?:[^']|'')+)'|([^!']+))!)?(.+)$/.exec(String(a1).trim());
  if (!m) return null;
  const sheetName = m[1] !== undefined ? m[1].replace(/''/g, '\'') : (m[2] !== undefined ? m[2] : null);
  const ref = m[3].replace(/\$/g, '');
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const parsePart = (p) => {
    const pm = /^([A-Za-z]{0,3})(\d*)$/.exec(p);
    if (!pm || (!pm[1] && !pm[2])) return null;
    return {col: pm[1] ? lettersToColumn(pm[1]) : null, row: pm[2] ? Number(pm[2]) : null};
  };
  const a = parsePart(parts[0]);
  if (!a) return null;
  const b = parts.length === 2 ? parsePart(parts[1]) : a;
  if (!b) return null;
  if (parts.length === 1 && (a.col === null || a.row === null)) return null;
  const r1 = a.row === null ? 1 : a.row;
  const c1 = a.col === null ? 1 : a.col;
  const r2 = b.row === null ? (a.row === null ? maxRows : Math.max(maxRows, r1)) : b.row;
  const c2 = b.col === null ? (a.col === null ? maxCols : Math.max(maxCols, c1)) : b.col;
  if (r1 < 1 || c1 < 1) return null;
  const row = Math.min(r1, r2);
  const col = Math.min(c1, c2);
  const open = a.row === null || a.col === null || b.row === null || b.col === null;
  return {
    sheetName,
    row,
    col,
    numRows: Math.abs(r2 - r1) + 1,
    numCols: Math.abs(c2 - c1) + 1,
    hint: open ? ref.toUpperCase() : null,
  };
}

/**
 * Compresses a map of cell -> value into rectangular A1 ranges whose cells
 * share the same (JSON-equal) value. Useful for readable dumps.
 * @param {!Array<{row: number, col: number, value: *}>} cells
 * @return {!Object<string, *>} A1 range -> value.
 */
function compressCells(cells) {
  const byRow = new Map();
  for (const cell of cells) {
    if (!byRow.has(cell.row)) byRow.set(cell.row, []);
    byRow.get(cell.row).push(cell);
  }
  const rows = [...byRow.keys()].sort((x, y) => x - y);
  /** @type {!Map<string, {row: number, col: number, endRow: number, endCol: number, value: *}>} */
  let open = new Map();
  const done = [];
  for (const row of rows) {
    const rowCells = byRow.get(row).sort((x, y) => x.col - y.col);
    const runs = [];
    for (const cell of rowCells) {
      const key = JSON.stringify(cell.value);
      const last = runs[runs.length - 1];
      if (last && last.key === key && last.endCol === cell.col - 1) {
        last.endCol = cell.col;
      } else {
        runs.push({col: cell.col, endCol: cell.col, key, value: cell.value});
      }
    }
    const nextOpen = new Map();
    for (const run of runs) {
      const id = `${run.col}:${run.endCol}:${run.key}`;
      const rect = open.get(id);
      if (rect && rect.endRow === row - 1) {
        rect.endRow = row;
        open.delete(id);
        nextOpen.set(id, rect);
      } else {
        nextOpen.set(id, {row, col: run.col, endRow: row, endCol: run.endCol, value: run.value});
      }
    }
    for (const rect of open.values()) done.push(rect);
    open = nextOpen;
  }
  for (const rect of open.values()) done.push(rect);
  done.sort((x, y) => x.row - y.row || x.col - y.col);
  const out = {};
  for (const rect of done) {
    out[rangeA1(rect.row, rect.col, rect.endRow - rect.row + 1, rect.endCol - rect.col + 1)] =
        rect.value;
  }
  return out;
}

module.exports = {columnToLetters, lettersToColumn, cellA1, rangeA1, parseA1, compressCells};
