/**
 * @fileoverview Host-side fixture helpers for the in-memory spreadsheet,
 * exposed as `runtime.spreadsheet`. Everything returned from here is plain
 * host-realm data, so it can be compared with assert.deepStrictEqual.
 */
'use strict';

const {cellA1, columnToLetters, compressCells, parseA1, rangeA1} = require('./a1');
const {isDate} = require('./realm');
const {SheetMock, displayOf, FORMAT_DEFAULTS} = require('./spreadsheetMock');
const B = require('./spreadsheetBuilders');

/**
 * Encodes a cell value for JSON fixtures (Dates -> {$date: iso}).
 * @param {*} v
 * @return {*}
 */
function encodeValue(v) {
  return isDate(v) ? {$date: v.toISOString()} : v;
}

/**
 * Decodes a fixture value ({$date: iso} -> Date; null/undefined -> '').
 * @param {*} v
 * @return {*}
 */
function decodeValue(v) {
  if (v === null || v === undefined) return '';
  if (v && typeof v === 'object' && !Array.isArray(v) && typeof v.$date === 'string') return new Date(v.$date);
  if (isDate(v)) return new Date(v.getTime());
  return v;
}

/**
 * Expands an A1 range key ('A1', 'B2:C5') into cell coordinates.
 * @param {string} a1
 * @param {!SheetMock} sheet
 * @return {!Array<{r: number, c: number}>}
 */
function expandA1(a1, sheet) {
  const p = parseA1(a1, sheet._maxRows, sheet._maxCols);
  if (!p) throw new Error(`Invalid A1 notation in fixture: ${a1}`);
  const out = [];
  for (let r = p.row; r < p.row + p.numRows; r++) {
    for (let c = p.col; c < p.col + p.numCols; c++) out.push({r, c});
  }
  return out;
}

/**
 * Fixture helpers bound to the runtime's active spreadsheet.
 */
class SpreadsheetHarness {
  /**
   * @param {!Object} env
   * @param {!Object} spreadsheet SpreadsheetMock (the one SpreadsheetApp.getActive() returns).
   * @param {!Object} app SpreadsheetAppMock.
   */
  constructor(env, spreadsheet, app) {
    /** @const */
    this.env = env;
    /** The Spreadsheet mock instance that the code under test sees. */
    this.active = spreadsheet;
    /** The SpreadsheetApp mock instance. */
    this.app = app;
  }

  /** @return {!Array<string>} Sheet names in tab order. */
  sheetNames() {
    return this.active._sheets.map((s) => s._name);
  }

  /**
   * @param {string} name
   * @return {!SheetMock} The sheet (throws when missing).
   */
  getSheet(name) {
    const sheet = this.active.getSheetByName(name);
    if (!sheet) throw new Error(`No sheet named "${name}" (have: ${this.sheetNames().join(', ')})`);
    return sheet;
  }

  /** @return {boolean} */
  hasSheet(name) {
    return Boolean(this.active.getSheetByName(name));
  }

  /**
   * Creates a sheet (or replaces an existing sheet's contents) with raw values.
   * Values are stored exactly as given (no auto-typing) unless
   * `options.autoType` is true, in which case they are written through
   * Range.setValues semantics (auto-typing, formulas).
   *
   * @param {string} name
   * @param {!Array<!Array<*>>=} rows Ragged rows are padded with ''.
   * @param {{autoType: (boolean|undefined), index: (number|undefined),
   *     frozenRows: (number|undefined), frozenColumns: (number|undefined),
   *     maxRows: (number|undefined), maxColumns: (number|undefined),
   *     activate: (boolean|undefined)}=} options
   * @return {!SheetMock}
   */
  addSheet(name, rows = [], options = {}) {
    let sheet = this.active.getSheetByName(name);
    if (sheet) {
      sheet._grid = [];
      sheet._touch();
    } else {
      sheet = this.active._createSheet(name, options.index === undefined ? this.active._sheets.length :
        options.index, Boolean(options.activate), {maxRows: options.maxRows, maxColumns: options.maxColumns});
    }
    const width = rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
    if (rows.length && width) {
      if (options.autoType) {
        const padded = rows.map((r) => Array.from({length: width}, (_, j) => (r && j < r.length ? r[j] : '')));
        sheet._writeBlock(1, 1, padded, 'value');
      } else {
        rows.forEach((row, i) => (row || []).forEach((v, j) => {
          const value = decodeValue(v);
          if (value === '') return;
          sheet._ensureCell(i + 1, j + 1).v = value;
        }));
      }
    }
    if (options.maxRows && sheet._maxRows < options.maxRows) sheet._maxRows = options.maxRows;
    if (options.maxColumns && sheet._maxCols < options.maxColumns) sheet._maxCols = options.maxColumns;
    if (options.frozenRows !== undefined) sheet._frozenRows = options.frozenRows;
    if (options.frozenColumns !== undefined) sheet._frozenCols = options.frozenColumns;
    sheet._touch();
    return sheet;
  }

  /**
   * Deletes a sheet if present.
   * @param {string} name
   */
  removeSheet(name) {
    const sheet = this.active.getSheetByName(name);
    if (!sheet) return;
    this.active._sheets = this.active._sheets.filter((s) => s !== sheet);
    sheet._deleted = true;
    if (this.active._activeSheet === sheet) this.active._activeSheet = this.active._sheets[0] || null;
  }

  /**
   * Values of the sheet's data range (rows 1..lastRow, cols 1..lastColumn).
   * @param {string} name
   * @return {!Array<!Array<*>>} [] when the sheet is empty.
   */
  getSheetValues(name) {
    const sheet = this.getSheet(name);
    const {lastRow, lastCol} = sheet._computeExtents();
    const out = [];
    for (let r = 1; r <= lastRow; r++) {
      const row = [];
      for (let c = 1; c <= lastCol; c++) {
        const cell = sheet._cell(r, c);
        const v = cell ? cell.v : '';
        row.push(isDate(v) ? new Date(v.getTime()) : v);
      }
      out.push(row);
    }
    return out;
  }

  /**
   * Display strings of the data range.
   * @param {string} name
   * @return {!Array<!Array<string>>}
   */
  getSheetDisplayValues(name) {
    const sheet = this.getSheet(name);
    const tz = this.active.getSpreadsheetTimeZone();
    const {lastRow, lastCol} = sheet._computeExtents();
    const out = [];
    for (let r = 1; r <= lastRow; r++) {
      const row = [];
      for (let c = 1; c <= lastCol; c++) {
        const cell = sheet._cell(r, c);
        row.push(cell ? displayOf(cell.v, cell.fmt, tz) : '');
      }
      out.push(row);
    }
    return out;
  }

  /**
   * Values of one row (1-based) up to the sheet's last column.
   * @param {string} name
   * @param {number} row
   * @return {!Array<*>}
   */
  getRowValues(name, row) {
    const sheet = this.getSheet(name);
    const lastCol = Math.max(1, sheet._computeExtents().lastCol);
    return Array.from({length: lastCol}, (_, j) => {
      const cell = sheet._cell(row, j + 1);
      return cell ? cell.v : '';
    });
  }

  /**
   * Everything known about one cell.
   * @param {string} name
   * @param {string} a1 e.g. 'B3'
   * @return {{value: *, display: string, formula: string, note: string,
   *     validation: ?Object, format: !Object, link: ?string}}
   */
  getCell(name, a1) {
    const sheet = this.getSheet(name);
    const p = parseA1(a1, sheet._maxRows, sheet._maxCols);
    if (!p) throw new Error(`Invalid A1 notation: ${a1}`);
    const cell = sheet._cell(p.row, p.col);
    return {
      value: cell ? cell.v : '',
      display: cell ? displayOf(cell.v, cell.fmt, this.active.getSpreadsheetTimeZone()) : '',
      formula: cell && cell.f ? cell.f : '',
      note: cell ? cell.note : '',
      validation: cell && cell.dv ? cell.dv.describe() : null,
      format: {...FORMAT_DEFAULTS, ...(cell && cell.fmt ? cell.fmt : {})},
      link: cell && cell.link ? cell.link : null,
    };
  }

  /**
   * Snapshot of a sheet: values, formulas, notes, validations, formats (only
   * non-default ones), layout. Per-cell maps are compressed into A1 ranges of
   * identical entries (e.g. {'C2:C10': {criteria: 'CHECKBOX', ...}}).
   * @param {string} name
   * @param {{compress: (boolean|undefined), display: (boolean|undefined)}=} options
   * @return {!Object}
   */
  dumpSheet(name, options = {}) {
    const sheet = this.getSheet(name);
    const compress = options.compress !== false;
    const collect = (pick) => {
      const cells = [];
      sheet._grid.forEach((row, i) => {
        if (!row) return;
        row.forEach((cell, j) => {
          if (!cell) return;
          const value = pick(cell);
          if (value !== undefined) cells.push({row: i + 1, col: j + 1, value});
        });
      });
      if (compress) return compressCells(cells);
      return Object.fromEntries(cells.map((c) => [cellA1(c.row, c.col), c.value]));
    };
    const dump = {
      name: sheet._name,
      sheetId: sheet._sheetId,
      index: sheet.getIndex(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      maxRows: sheet._maxRows,
      maxColumns: sheet._maxCols,
      frozenRows: sheet._frozenRows,
      frozenColumns: sheet._frozenCols,
      hidden: sheet._hidden,
      tabColor: sheet._tabColor,
      values: this.getSheetValues(name),
      formulas: collect((c) => (c.f ? c.f : undefined)),
      notes: collect((c) => (c.note ? c.note : undefined)),
      validations: collect((c) => (c.dv ? c.dv.describe() : undefined)),
      formats: collect((c) => (c.fmt && Object.keys(c.fmt).length ? {...c.fmt} : undefined)),
      links: collect((c) => (c.link && !c.f ? c.link : undefined)),
      hiddenRows: [...sheet._hiddenRows].sort((a, b) => a - b),
      hiddenColumns: [...sheet._hiddenCols].sort((a, b) => a - b),
      columnWidths: Object.fromEntries([...sheet._colWidths].sort((a, b) => a[0] - b[0])),
      rowHeights: Object.fromEntries([...sheet._rowHeights].sort((a, b) => a[0] - b[0])),
      merges: sheet._merges.map((m) => rangeA1(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1)),
      borders: sheet._borders.map((b) => ({...b})),
      filter: sheet._filter ? sheet._filter.describe() : null,
      conditionalFormatRules: sheet._cfRules.map((r) => (r.describe ? r.describe() : String(r))),
      protections: sheet._protections.map((p) => p.describe()),
      bandings: sheet._bandings.map((b) => b.describe()),
    };
    if (options.display) dump.displayValues = this.getSheetDisplayValues(name);
    return dump;
  }

  /**
   * Renders the data range as an aligned text table (handy for debugging).
   * @param {string} name
   * @param {{maxWidth: (number|undefined), display: (boolean|undefined)}=} options
   * @return {string}
   */
  printSheet(name, options = {}) {
    const maxWidth = options.maxWidth || 24;
    const rows = options.display === false ? this.getSheetValues(name).map((r) => r.map((v) => JSON.stringify(v))) :
      this.getSheetDisplayValues(name);
    if (rows.length === 0) return `(sheet "${name}" is empty)`;
    const cols = rows[0].length;
    const clip = (s) => {
      const t = String(s).replace(/\n/g, '\u21b5');
      return t.length > maxWidth ? `${t.slice(0, maxWidth - 1)}\u2026` : t;
    };
    const header = ['', ...Array.from({length: cols}, (_, j) => columnToLetters(j + 1))];
    const body = rows.map((r, i) => [String(i + 1), ...r.map(clip)]);
    const widths = header.map((h, j) => Math.max(h.length, ...body.map((r) => r[j].length)));
    const line = (cells) => cells.map((c, j) => c.padEnd(widths[j])).join(' | ');
    return [line(header), widths.map((w) => '-'.repeat(w)).join('-+-'), ...body.map(line)].join('\n');
  }

  /**
   * Serializes the whole spreadsheet to a JSON-compatible fixture.
   * @return {!Object}
   */
  toJSON() {
    return {
      id: this.active._id,
      name: this.active._name,
      timeZone: this.active._timeZone,
      locale: this.active._locale,
      activeSheet: this.active.getActiveSheet() ? this.active.getActiveSheet()._name : null,
      sheets: this.active._sheets.map((sheet) => {
        const d = this.dumpSheet(sheet._name);
        return {
          name: d.name,
          sheetId: d.sheetId,
          maxRows: d.maxRows,
          maxColumns: d.maxColumns,
          frozenRows: d.frozenRows,
          frozenColumns: d.frozenColumns,
          hidden: d.hidden,
          tabColor: d.tabColor,
          values: d.values.map((row) => row.map(encodeValue)),
          formulas: d.formulas,
          notes: d.notes,
          validations: d.validations,
          formats: d.formats,
          links: d.links,
          hiddenRows: d.hiddenRows,
          hiddenColumns: d.hiddenColumns,
          columnWidths: d.columnWidths,
          rowHeights: d.rowHeights,
          merges: d.merges,
        };
      }),
    };
  }

  /**
   * Replaces all sheets with the fixture's content (in place, so references
   * captured by the code under test - e.g. a top-level
   * `const ss = SpreadsheetApp.getActive()` - stay valid).
   *
   * Fixture shape (all sheet fields optional except name):
   * {id, name, timeZone, locale, activeSheet, sheets: [{name, values: [[...]],
   *  formulas: {A1: '=...'}, notes: {A1: 'x'}, validations: {'C2:C9':
   *  {criteria: 'CHECKBOX', values: [], allowInvalid: true, helpText: ''}},
   *  formats: {A1: {background: '#ff0000'}}, links: {A1: url}, frozenRows,
   *  frozenColumns, hidden, tabColor, maxRows, maxColumns, hiddenRows,
   *  hiddenColumns, columnWidths: {1: 200}, rowHeights: {}, merges: ['A1:B1']}]}
   * Values are stored raw; {$date: iso} decodes to a Date. VALUE_IN_RANGE
   * criteria reference ranges as {$range: 'Sheet!A1:A9'}.
   *
   * @param {!Object|string} fixture Object or JSON string.
   * @return {!SpreadsheetHarness}
   */
  loadFixture(fixture) {
    const json = typeof fixture === 'string' ? JSON.parse(fixture) : fixture;
    const ss = this.active;
    for (const s of ss._sheets) s._deleted = true;
    ss._sheets = [];
    ss._activeSheet = null;
    ss._namedRanges.clear();
    if (json.id) {
      this.env.spreadsheets.delete(ss._id);
      ss._id = String(json.id);
      this.env.spreadsheets.set(ss._id, ss);
    }
    if (json.name) ss._name = String(json.name);
    if (json.timeZone) ss._timeZone = String(json.timeZone);
    if (json.locale) ss._locale = String(json.locale);
    const pendingValidations = [];
    for (const spec of json.sheets || []) {
      const sheet = this.addSheet(spec.name, spec.values || [], {
        maxRows: spec.maxRows, maxColumns: spec.maxColumns,
        frozenRows: spec.frozenRows, frozenColumns: spec.frozenColumns,
      });
      if (spec.sheetId !== undefined) sheet._sheetId = spec.sheetId;
      if (spec.maxRows) sheet._maxRows = Math.max(spec.maxRows, sheet._grid.length);
      if (spec.maxColumns) sheet._maxCols = spec.maxColumns;
      sheet._hidden = Boolean(spec.hidden);
      sheet._tabColor = spec.tabColor ? B.normalizeColor(spec.tabColor) : null;
      for (const [a1, formula] of Object.entries(spec.formulas || {})) {
        for (const {r, c} of expandA1(a1, sheet)) sheet._ensureCell(r, c).f = String(formula);
      }
      for (const [a1, note] of Object.entries(spec.notes || {})) {
        for (const {r, c} of expandA1(a1, sheet)) sheet._ensureCell(r, c).note = String(note);
      }
      for (const [a1, fmt] of Object.entries(spec.formats || {})) {
        for (const {r, c} of expandA1(a1, sheet)) {
          const cell = sheet._ensureCell(r, c);
          cell.fmt = {...(cell.fmt || {}), ...fmt};
        }
      }
      for (const [a1, url] of Object.entries(spec.links || {})) {
        for (const {r, c} of expandA1(a1, sheet)) sheet._ensureCell(r, c).link = String(url);
      }
      for (const [a1, desc] of Object.entries(spec.validations || {})) pendingValidations.push({sheet, a1, desc});
      (spec.hiddenRows || []).forEach((r) => sheet._hiddenRows.add(Number(r)));
      (spec.hiddenColumns || []).forEach((c) => sheet._hiddenCols.add(Number(c)));
      Object.entries(spec.columnWidths || {}).forEach(([c, w]) => sheet._colWidths.set(Number(c), w));
      Object.entries(spec.rowHeights || {}).forEach(([r, h]) => sheet._rowHeights.set(Number(r), h));
      for (const m of spec.merges || []) {
        const p = parseA1(m, sheet._maxRows, sheet._maxCols);
        if (p) sheet._merges.push({r1: p.row, c1: p.col, r2: p.row + p.numRows - 1, c2: p.col + p.numCols - 1});
      }
      sheet._touch();
    }
    for (const {sheet, a1, desc} of pendingValidations) {
      const values = (desc.values || []).map((v) => {
        if (v && typeof v === 'object' && typeof v.$range === 'string') return ss.getRange(v.$range);
        return v;
      });
      const rule = new B.DataValidationBuilderMock(this.env.realm)
          .withCriteria(desc.criteria, [])
          .setAllowInvalid(desc.allowInvalid !== false)
          .setHelpText(desc.helpText || '')
          .build();
      rule._values = values;
      for (const {r, c} of expandA1(a1, sheet)) sheet._ensureCell(r, c).dv = rule;
    }
    if (ss._sheets.length === 0) ss._createSheet('Sheet1', 0, true);
    const active = json.activeSheet ? ss.getSheetByName(json.activeSheet) : null;
    ss._activeSheet = active || ss._sheets[0];
    return this;
  }
}

module.exports = {SpreadsheetHarness, encodeValue, decodeValue};
