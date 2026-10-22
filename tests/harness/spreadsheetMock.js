/**
 * @fileoverview In-memory mock of SpreadsheetApp (Spreadsheet, Sheet, Range).
 *
 * Fidelity notes (see tests/README.md for the full list):
 *  - Indices are 1-based like Apps Script; empty cells read as ''.
 *  - Strings written with setValue(s) are auto-typed like the Sheets UI does
 *    ('123' -> 123, '1,234' -> 1234, '5%' -> 0.05, 'true' -> true), unless the
 *    cell's number format is plain text ('@'). Dates are NOT parsed. A leading
 *    apostrophe forces text. Strings starting with '=' become formulas; only
 *    trivial formulas (HYPERLINK, literals) are evaluated.
 *  - Checkbox cells store booleans; clearing them makes them unchecked (false).
 *  - Writing beyond getMaxRows()/getMaxColumns() grows the sheet (pass
 *    `strictBounds: true` to throw like older Sheets did instead).
 *  - Formats, notes, validations are stored per cell so tests can assert them.
 */
'use strict';

const {defineHidden} = require('./uiMock');
const {cellA1, rangeA1, parseA1} = require('./a1');
const {isDate} = require('./realm');
const {formatDateSheets, formatNumberSheets} = require('./dateFormat');
const B = require('./spreadsheetBuilders');

const DEFAULT_MAX_ROWS = 1000;
const DEFAULT_MAX_COLS = 26;
const DEFAULT_COLUMN_WIDTH = 100;
const DEFAULT_ROW_HEIGHT = 21;
const DEFAULT_NUMBER_FORMAT = '0.###############';

/** Values reported by format getters for cells without explicit formatting. */
const FORMAT_DEFAULTS = Object.freeze({
  background: '#ffffff',
  fontColor: '#000000',
  fontWeight: 'normal',
  fontStyle: 'normal',
  fontSize: 10,
  fontFamily: 'Arial',
  fontLine: 'none',
  horizontalAlignment: 'general',
  verticalAlignment: 'bottom',
  wrapStrategy: 'OVERFLOW',
  numberFormat: DEFAULT_NUMBER_FORMAT,
  textDirection: null,
  textRotation: 0,
  showHyperlink: true,
});

const NUMBER_RE = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
const GROUPED_NUMBER_RE = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/;
const PERCENT_RE = /^([+-]?(\d+\.?\d*|\.\d+))%$/;

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

/** @return {!Object} A fresh cell record. */
function newCell() {
  return {v: '', f: null, note: '', dv: null, fmt: null, link: null, rich: null};
}

/**
 * @param {?Object} cell
 * @return {?Object} Deep-enough copy of a cell record.
 */
function cloneCell(cell) {
  if (!cell) return undefined;
  return {...cell, v: isDate(cell.v) ? new Date(cell.v.getTime()) : cell.v,
    fmt: cell.fmt ? {...cell.fmt} : null};
}

/**
 * A cell counts as content (for getLastRow etc.) when it has a value or formula.
 * @param {?Object} cell
 * @return {boolean}
 */
function isBlankCell(cell) {
  return !cell || ((cell.v === '' || cell.v === null || cell.v === undefined) && !cell.f);
}

/**
 * Apps Script-style type name used in "parameters don't match" errors.
 * @param {*} v
 * @return {string}
 */
function typeName(v) {
  if (v === null || v === undefined) return 'null';
  if (Array.isArray(v)) {
    if (v.length && Array.isArray(v[0])) return `${v[0].length ? typeName(v[0][0]) : 'Object'}[][]`;
    return `${v.length ? typeName(v[0]) : 'Object'}[]`;
  }
  if (isDate(v)) return 'Date';
  switch (typeof v) {
    case 'string': return 'String';
    case 'number': return 'number';
    case 'boolean': return 'Boolean';
    case 'function': return 'Function';
    default: return 'Object';
  }
}

/**
 * Applies Sheets-like auto-typing to a string typed into a cell.
 * @param {string} str
 * @return {{value: *, numberFormat: (string|undefined)}}
 */
function autoType(str) {
  const t = str.trim();
  if (t === '') return {value: str};
  if (NUMBER_RE.test(t)) return {value: Number(t)};
  if (GROUPED_NUMBER_RE.test(t)) return {value: Number(t.replace(/,/g, ''))};
  const pct = PERCENT_RE.exec(t);
  if (pct) return {value: Number(pct[1]) / 100, numberFormat: pct[1].includes('.') ? '0.00%' : '0%'};
  if (/^(true|false)$/i.test(t)) return {value: t.toLowerCase() === 'true'};
  return {value: str};
}

/**
 * Evaluates the handful of formulas the harness understands.
 * @param {string} formula Including the leading '='.
 * @param {?function(string, !Object): *} customEvaluator
 * @param {!Object} where {sheet, row, column}
 * @return {{value: *, link: (string|undefined), unevaluated: (boolean|undefined)}}
 */
function evaluateFormula(formula, customEvaluator, where) {
  if (customEvaluator) {
    const custom = customEvaluator(formula, where);
    if (custom !== undefined) return {value: custom};
  }
  const body = formula.slice(1).trim();
  const unq = (s) => s.replace(/""/g, '"');
  const link = /^HYPERLINK\(\s*"((?:[^"]|"")*)"\s*(?:[,;]\s*"((?:[^"]|"")*)"\s*)?\)$/i.exec(body);
  if (link) return {value: unq(link[2] !== undefined ? link[2] : link[1]), link: unq(link[1])};
  if (/^HYPERLINK\(/i.test(body)) return {value: '#ERROR!', unevaluated: true};
  const str = /^"((?:[^"]|"")*)"$/.exec(body);
  if (str) return {value: unq(str[1])};
  if (NUMBER_RE.test(body)) return {value: Number(body)};
  if (/^(TRUE|FALSE)(\(\))?$/i.test(body)) return {value: /^TRUE/i.test(body)};
  return {value: '', unevaluated: true};
}

/**
 * Formats a stored value the way Sheets displays it.
 * @param {*} value
 * @param {?Object} fmt Cell format record.
 * @param {string} timeZone Spreadsheet time zone.
 * @return {string}
 */
function displayOf(value, fmt, timeZone) {
  if (value === '' || value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') return formatNumberSheets(value, fmt && fmt.numberFormat);
  if (isDate(value)) return formatDateSheets(value, timeZone, fmt && fmt.numberFormat);
  return String(value);
}

/**
 * Inserts `n` holes into `arr` at `index`, shifting the tail right.
 * Avoids spreading huge argument lists into splice().
 * @param {!Array} arr
 * @param {number} index 0-based.
 * @param {number} n
 */
function insertHoles(arr, index, n) {
  if (arr.length <= index) return;
  const tail = arr.splice(index);
  arr.length = index + n;
  for (let i = 0; i < tail.length; i++) arr[index + n + i] = tail[i];
}

/**
 * Shifts integer keys of a Set or Map after inserting/deleting lines.
 * @param {!Set<number>|!Map<number, *>} coll
 * @param {number} at First affected 1-based index.
 * @param {number} delta Positive for inserts, negative for deletes.
 * @return {!Set<number>|!Map<number, *>} A new collection.
 */
function shiftKeys(coll, at, delta) {
  const isMap = coll instanceof Map;
  const out = isMap ? new Map() : new Set();
  for (const entry of coll) {
    const key = isMap ? entry[0] : entry;
    let newKey = key;
    if (delta < 0 && key >= at && key < at - delta) continue;  // deleted
    if (key >= at) newKey = key + delta;
    if (isMap) out.set(newKey, entry[1]);
    else out.add(newKey);
  }
  return out;
}

/**
 * @param {{r1: number, c1: number, r2: number, c2: number}} a
 * @param {{r1: number, c1: number, r2: number, c2: number}} b
 * @return {boolean}
 */
function intersects(a, b) {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

/**
 * @param {{r1: number, c1: number, r2: number, c2: number}} outer
 * @param {{r1: number, c1: number, r2: number, c2: number}} inner
 * @return {boolean}
 */
function contains(outer, inner) {
  return outer.r1 <= inner.r1 && outer.r2 >= inner.r2 && outer.c1 <= inner.c1 && outer.c2 >= inner.c2;
}

/**
 * Sheets sort order: numbers/dates < text < booleans; blanks always last.
 * @param {*} a
 * @param {*} b
 * @param {boolean} ascending
 * @return {number}
 */
function compareCellValues(a, b, ascending) {
  const rank = (v) => {
    if (v === '' || v === null || v === undefined) return 4;
    if (typeof v === 'number' || isDate(v)) return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 3;
  };
  const ra = rank(a);
  const rb = rank(b);
  if (ra === 4 || rb === 4) return ra === rb ? 0 : (ra === 4 ? 1 : -1);
  let result;
  if (ra !== rb) {
    result = ra - rb;
  } else if (ra === 0) {
    result = Number(a) - Number(b);
  } else if (ra === 1) {
    result = a.localeCompare(b, 'en', {sensitivity: 'base'});
  } else {
    result = Number(a) - Number(b);
  }
  return ascending ? result : -result;
}

// ---------------------------------------------------------------------------
// Filter & banding (need Range)
// ---------------------------------------------------------------------------

/** Filter attached to a sheet (Range.createFilter()). */
class FilterMock {
  /** @param {!RangeMock} range */
  constructor(range) {
    defineHidden(this, {_range: range, _criteria: new Map()});
  }

  getRange() {
    return this._range;
  }

  remove() {
    if (this._range._sheet._filter === this) this._range._sheet._filter = null;
  }

  setColumnFilterCriteria(columnPosition, filterCriteria) {
    if (filterCriteria === null) this._criteria.delete(columnPosition);
    else this._criteria.set(columnPosition, filterCriteria);
    return this;
  }

  getColumnFilterCriteria(columnPosition) {
    return this._criteria.get(columnPosition) || null;
  }

  removeColumnFilterCriteria(columnPosition) {
    this._criteria.delete(columnPosition);
    return this;
  }

  sort(columnPosition, ascending) {
    const r = this._range;
    if (r._numRows > 1) {
      r._sheet.getRange(r._row + 1, r._col, r._numRows - 1, r._numCols)
          .sort({column: columnPosition, ascending: ascending !== false});
    }
    return this;
  }

  describe() {
    return {
      range: this._range.getA1Notation(),
      criteria: Object.fromEntries([...this._criteria].map(([k, v]) => [k, v.describe ? v.describe() : v])),
    };
  }

  toString() {
    return 'Filter';
  }
}

/** Banding (recorded only). */
class BandingMock {
  constructor(range, theme, kind) {
    defineHidden(this, {_range: range, _theme: theme, _kind: kind});
  }

  getRange() {
    return this._range;
  }

  remove() {
    const sheet = this._range._sheet;
    sheet._bandings = sheet._bandings.filter((b) => b !== this);
  }

  setHeaderRowColor() {
    return this;
  }

  setFirstRowColor() {
    return this;
  }

  setSecondRowColor() {
    return this;
  }

  setFooterRowColor() {
    return this;
  }

  describe() {
    return {range: this._range.getA1Notation(), theme: this._theme, kind: this._kind};
  }

  toString() {
    return 'Banding';
  }
}

// ---------------------------------------------------------------------------
// Range
// ---------------------------------------------------------------------------

/**
 * Range mock. Coordinates are fixed at creation (like Apps Script ranges).
 */
class RangeMock {
  /**
   * @param {!SheetMock} sheet
   * @param {number} row
   * @param {number} col
   * @param {number} numRows
   * @param {number} numCols
   * @param {?string=} hint A1 notation to report for open-ended ranges ('A:A').
   */
  constructor(sheet, row, col, numRows, numCols, hint = null) {
    defineHidden(this, {
      _sheet: sheet, _row: row, _col: col, _numRows: numRows, _numCols: numCols, _hint: hint,
    });
  }

  /** @return {!Object} realm helpers */
  get _realm() {
    return this._sheet._env.realm;
  }

  /** @return {{r1: number, c1: number, r2: number, c2: number}} */
  _rect() {
    return {
      r1: this._row, c1: this._col,
      r2: this._row + this._numRows - 1, c2: this._col + this._numCols - 1,
    };
  }

  /**
   * Calls fn(r, c, i, j) for every cell (row-major).
   * @param {function(number, number, number, number)} fn
   */
  _each(fn) {
    for (let i = 0; i < this._numRows; i++) {
      for (let j = 0; j < this._numCols; j++) fn(this._row + i, this._col + j, i, j);
    }
  }

  /**
   * Builds a host 2D array from fn(cellOrUndefined, r, c).
   * @param {function(?Object, number, number): *} fn
   * @return {!Array<!Array<*>>}
   */
  _map(fn) {
    const out = [];
    for (let i = 0; i < this._numRows; i++) {
      const row = [];
      for (let j = 0; j < this._numCols; j++) {
        const r = this._row + i;
        const c = this._col + j;
        row.push(fn(this._sheet._cell(r, c), r, c));
      }
      out.push(row);
    }
    return out;
  }

  /**
   * @param {function(?Object, number, number): *} fn
   * @return {*} fn applied to the top-left cell.
   */
  _topLeft(fn) {
    return fn(this._sheet._cell(this._row, this._col), this._row, this._col);
  }

  /**
   * Validates a 2D array argument exactly like Apps Script.
   * @param {*} values
   * @param {string} method e.g. 'Range.setValues'
   */
  _check2D(values, method) {
    const realm = this._realm;
    if (!Array.isArray(values) || values.some((row) => !Array.isArray(row))) {
      throw realm.error(`The parameters (${typeName(values)}) don't match the method signature ` +
          `for SpreadsheetApp.${method}.`);
    }
    if (values.length !== this._numRows) {
      throw realm.error('The number of rows in the data does not match the number of rows in the ' +
          `range. The data has ${values.length} but the range has ${this._numRows}.`);
    }
    for (const row of values) {
      if (row.length !== this._numCols) {
        throw realm.error('The number of columns in the data does not match the number of columns ' +
            `in the range. The data has ${row.length} but the range has ${this._numCols}.`);
      }
    }
  }

  /** @return {!Array<!Array<*>>} A 2D array filled with `value`. */
  _fill(value) {
    return Array.from({length: this._numRows}, () => new Array(this._numCols).fill(value));
  }

  // --- geometry -------------------------------------------------------------

  getRow() {
    return this._row;
  }

  getRowIndex() {
    return this._row;
  }

  getColumn() {
    return this._col;
  }

  getNumRows() {
    return this._numRows;
  }

  getNumColumns() {
    return this._numCols;
  }

  getHeight() {
    return this._numRows;
  }

  getWidth() {
    return this._numCols;
  }

  getLastRow() {
    return this._row + this._numRows - 1;
  }

  getLastColumn() {
    return this._col + this._numCols - 1;
  }

  getA1Notation() {
    return this._hint || rangeA1(this._row, this._col, this._numRows, this._numCols);
  }

  getSheet() {
    return this._sheet;
  }

  getGridId() {
    return this._sheet.getSheetId();
  }

  /**
   * @param {number} rowOffset
   * @param {number} columnOffset
   * @param {number=} numRows
   * @param {number=} numColumns
   * @return {!RangeMock}
   */
  offset(rowOffset, columnOffset, numRows, numColumns) {
    return this._sheet.getRange(this._row + rowOffset, this._col + columnOffset,
        numRows === undefined ? this._numRows : numRows,
        numColumns === undefined ? this._numCols : numColumns);
  }

  /**
   * @param {number} row 1-based, relative to this range.
   * @param {number} column 1-based, relative to this range.
   * @return {!RangeMock}
   */
  getCell(row, column) {
    if (row < 1 || column < 1 || row > this._numRows || column > this._numCols) {
      throw this._realm.error('Cell reference out of range');
    }
    return this._sheet.getRange(this._row + row - 1, this._col + column - 1);
  }

  /** @return {!RangeMock} */
  getDataRegion() {
    return this._sheet.getDataRange();
  }

  /** @return {!RangeMock} */
  activate() {
    this._sheet._ss._activeSheet = this._sheet;
    this._sheet._activeRange = this;
    return this;
  }

  /** @return {!RangeMock} */
  activateAsCurrentCell() {
    return this.activate();
  }

  // --- values ---------------------------------------------------------------

  getValues() {
    return this._realm.toContext(this._map((cell) => (cell ? cell.v : '')));
  }

  getValue() {
    return this._realm.toContext(this._topLeft((cell) => (cell ? cell.v : '')));
  }

  /**
   * @param {!Array<!Array<*>>} values
   * @return {!RangeMock}
   */
  setValues(values) {
    this._check2D(values, 'Range.setValues');
    this._sheet._writeBlock(this._row, this._col, values, 'value');
    return this;
  }

  /**
   * Sets every cell in the range to `value`.
   * @param {*} value
   * @return {!RangeMock}
   */
  setValue(value) {
    this._sheet._writeBlock(this._row, this._col, this._fill(value), 'value');
    return this;
  }

  getDisplayValues() {
    const tz = this._sheet._ss.getSpreadsheetTimeZone();
    return this._realm.toContext(this._map((cell) => (cell ? displayOf(cell.v, cell.fmt, tz) : '')));
  }

  getDisplayValue() {
    const tz = this._sheet._ss.getSpreadsheetTimeZone();
    return this._topLeft((cell) => (cell ? displayOf(cell.v, cell.fmt, tz) : ''));
  }

  getFormulas() {
    return this._realm.toContext(this._map((cell) => (cell && cell.f ? cell.f : '')));
  }

  getFormula() {
    return this._topLeft((cell) => (cell && cell.f ? cell.f : ''));
  }

  setFormula(formula) {
    this._sheet._writeBlock(this._row, this._col, this._fill(formula), 'formula');
    return this;
  }

  setFormulas(formulas) {
    this._check2D(formulas, 'Range.setFormulas');
    this._sheet._writeBlock(this._row, this._col, formulas, 'formula');
    return this;
  }

  getFormulaR1C1() {
    return this.getFormula();
  }

  getFormulasR1C1() {
    return this.getFormulas();
  }

  setFormulaR1C1(formula) {
    return this.setFormula(formula);
  }

  setFormulasR1C1(formulas) {
    return this.setFormulas(formulas);
  }

  /** @return {boolean} */
  isBlank() {
    let blank = true;
    this._each((r, c) => {
      if (!isBlankCell(this._sheet._cell(r, c))) blank = false;
    });
    return blank;
  }

  // --- rich text ------------------------------------------------------------

  setRichTextValue(value) {
    this._sheet._writeRich(this._row, this._col, this._fill(value));
    return this;
  }

  setRichTextValues(values) {
    this._check2D(values, 'Range.setRichTextValues');
    this._sheet._writeRich(this._row, this._col, values);
    return this;
  }

  getRichTextValue() {
    return this._topLeft((cell) => this._sheet._richOf(cell));
  }

  getRichTextValues() {
    return this._realm.toContext(this._map((cell) => this._sheet._richOf(cell)));
  }

  // --- clearing -------------------------------------------------------------

  /**
   * clear() | clear({contentsOnly, formatOnly, commentsOnly, validationsOnly})
   * @param {!Object=} options
   * @return {!RangeMock}
   */
  clear(options) {
    this._sheet._clearRect(this._rect(), options);
    return this;
  }

  clearContent() {
    this._sheet._clearRect(this._rect(), {contentsOnly: true});
    return this;
  }

  clearFormat() {
    this._sheet._clearRect(this._rect(), {formatOnly: true});
    return this;
  }

  clearNote() {
    this._sheet._clearRect(this._rect(), {commentsOnly: true});
    return this;
  }

  clearDataValidations() {
    this._sheet._clearRect(this._rect(), {validationsOnly: true});
    return this;
  }

  /** Trims and collapses whitespace in text cells. */
  trimWhitespace() {
    this._each((r, c) => {
      const cell = this._sheet._cell(r, c);
      if (cell && typeof cell.v === 'string' && !cell.f) {
        cell.v = cell.v.trim().replace(/\s+/g, ' ');
      }
    });
    this._sheet._touch();
    return this;
  }

  // --- notes ----------------------------------------------------------------

  setNote(note) {
    const text = note === null || note === undefined ? '' : String(note);
    this._each((r, c) => {
      if (text || this._sheet._cell(r, c)) this._sheet._ensureCell(r, c).note = text;
    });
    return this;
  }

  setNotes(notes) {
    this._check2D(notes, 'Range.setNotes');
    this._each((r, c, i, j) => {
      const n = notes[i][j];
      const text = n === null || n === undefined ? '' : String(n);
      if (text || this._sheet._cell(r, c)) this._sheet._ensureCell(r, c).note = text;
    });
    return this;
  }

  getNote() {
    return this._topLeft((cell) => (cell ? cell.note : ''));
  }

  getNotes() {
    return this._realm.toContext(this._map((cell) => (cell ? cell.note : '')));
  }

  // --- data validation ------------------------------------------------------

  /**
   * @param {?B.DataValidationMock} rule null clears validation.
   * @return {!RangeMock}
   */
  setDataValidation(rule) {
    if (rule !== null && !(rule instanceof B.DataValidationMock)) {
      throw this._realm.error(`The parameters (${typeName(rule)}) don't match the method signature ` +
          'for SpreadsheetApp.Range.setDataValidation.');
    }
    this._each((r, c) => this._sheet._applyValidation(r, c, rule));
    this._sheet._touch();
    return this;
  }

  setDataValidations(rules) {
    this._check2D(rules, 'Range.setDataValidations');
    this._each((r, c, i, j) => this._sheet._applyValidation(r, c, rules[i][j] || null));
    this._sheet._touch();
    return this;
  }

  getDataValidation() {
    return this._topLeft((cell) => (cell && cell.dv ? cell.dv : null));
  }

  getDataValidations() {
    return this._realm.toContext(this._map((cell) => (cell && cell.dv ? cell.dv : null)));
  }

  /**
   * insertCheckboxes() | insertCheckboxes(checked) | insertCheckboxes(checked, unchecked)
   * Sets every cell to the unchecked value (false), as documented.
   * @param {...*} args
   * @return {!RangeMock}
   */
  insertCheckboxes(...args) {
    const builder = new B.DataValidationBuilderMock(this._realm).requireCheckbox(...args.slice(0, 2));
    const rule = builder.build();
    const unchecked = rule.uncheckedValue();
    this._each((r, c) => {
      const cell = this._sheet._ensureCell(r, c);
      cell.dv = rule;
      cell.v = unchecked;
      cell.f = null;
      cell.rich = null;
      cell.link = null;
    });
    this._sheet._touch();
    return this;
  }

  /**
   * Removes checkbox validation; clears values equal to the checked/unchecked value.
   * @return {!RangeMock}
   */
  removeCheckboxes() {
    this._each((r, c) => {
      const cell = this._sheet._cell(r, c);
      if (!cell || !cell.dv || cell.dv.getCriteriaType() !== 'CHECKBOX') return;
      const {dv} = cell;
      if (cell.v === dv.checkedValue() || cell.v === dv.uncheckedValue()) cell.v = '';
      cell.dv = null;
    });
    this._sheet._touch();
    return this;
  }

  /** @return {!RangeMock} */
  check() {
    return this._setChecked(true);
  }

  /** @return {!RangeMock} */
  uncheck() {
    return this._setChecked(false);
  }

  /**
   * @param {boolean} checked
   * @return {!RangeMock}
   */
  _setChecked(checked) {
    this._each((r, c) => {
      const cell = this._sheet._cell(r, c);
      if (!cell || !cell.dv || cell.dv.getCriteriaType() !== 'CHECKBOX') return;
      const {dv} = cell;
      if (cell.v === dv.checkedValue() || cell.v === dv.uncheckedValue() || cell.v === '') {
        cell.v = checked ? dv.checkedValue() : dv.uncheckedValue();
      }
    });
    this._sheet._touch();
    return this;
  }

  /** @return {?boolean} true if all checked, false if all unchecked, else null. */
  isChecked() {
    let checked = 0;
    let unchecked = 0;
    let other = 0;
    this._each((r, c) => {
      const cell = this._sheet._cell(r, c);
      if (!cell || !cell.dv || cell.dv.getCriteriaType() !== 'CHECKBOX') {
        other++;
      } else if (cell.v === cell.dv.checkedValue()) {
        checked++;
      } else {
        unchecked++;
      }
    });
    if (other > 0) return null;
    if (checked && !unchecked) return true;
    if (unchecked && !checked) return false;
    return null;
  }

  // --- formatting (generated setters/getters are installed below) -----------

  /**
   * @param {string} key Format key.
   * @param {*} value Normalized value; null resets to default.
   */
  _setFormat(key, value) {
    this._each((r, c) => this._sheet._setCellFormat(r, c, key, value));
  }

  /**
   * @param {string} key
   * @param {*} cell
   * @return {*} Stored or default format value.
   */
  static _formatOf(key, cell) {
    if (cell && cell.fmt && cell.fmt[key] !== undefined) return cell.fmt[key];
    return FORMAT_DEFAULTS[key];
  }

  setWrap(isWrapEnabled) {
    this._setFormat('wrapStrategy', isWrapEnabled ? 'WRAP' : 'OVERFLOW');
    return this;
  }

  setWraps(values) {
    this._check2D(values, 'Range.setWraps');
    this._each((r, c, i, j) => this._sheet._setCellFormat(r, c, 'wrapStrategy',
        values[i][j] ? 'WRAP' : 'OVERFLOW'));
    return this;
  }

  getWrap() {
    return this._topLeft((cell) => !(cell && cell.fmt && cell.fmt.wrapStrategy &&
        cell.fmt.wrapStrategy !== 'WRAP'));
  }

  getWraps() {
    return this._realm.toContext(this._map((cell) => !(cell && cell.fmt && cell.fmt.wrapStrategy &&
        cell.fmt.wrapStrategy !== 'WRAP')));
  }

  setBackgroundRGB(red, green, blue) {
    return this.setBackground(`rgb(${red},${green},${blue})`);
  }

  setBackgroundObject(color) {
    return this.setBackground(color && color.asHexString ? color.asHexString() : color);
  }

  getBackgroundObject() {
    const hex = this.getBackground();
    return {asRgbColor: () => ({asHexString: () => hex}), asHexString: () => hex,
      getColorType: () => 'RGB', toString: () => 'Color'};
  }

  getFontColorObject() {
    const hex = this.getFontColor();
    return {asRgbColor: () => ({asHexString: () => hex}), asHexString: () => hex,
      getColorType: () => 'RGB', toString: () => 'Color'};
  }

  setFontColorObject(color) {
    return this.setFontColor(color && color.asHexString ? color.asHexString() : color);
  }

  /**
   * @param {?B.TextStyleMock} style
   * @return {!RangeMock}
   */
  setTextStyle(style) {
    this._each((r, c) => this._sheet._applyTextStyle(r, c, style));
    return this;
  }

  setTextStyles(styles) {
    this._check2D(styles, 'Range.setTextStyles');
    this._each((r, c, i, j) => this._sheet._applyTextStyle(r, c, styles[i][j]));
    return this;
  }

  getTextStyle() {
    return this._topLeft((cell) => this._sheet._textStyleOf(cell));
  }

  getTextStyles() {
    return this._realm.toContext(this._map((cell) => this._sheet._textStyleOf(cell)));
  }

  setShowHyperlink(showHyperlink) {
    this._setFormat('showHyperlink', Boolean(showHyperlink));
    return this;
  }

  setTextRotation(degrees) {
    this._setFormat('textRotation', typeof degrees === 'number' ? degrees : 0);
    return this;
  }

  setVerticalText(isVertical) {
    this._setFormat('verticalText', Boolean(isVertical));
    return this;
  }

  /**
   * Records a border call. Signature: (top, left, bottom, right, vertical,
   * horizontal[, color, style]).
   * @param {...*} args
   * @return {!RangeMock}
   */
  setBorder(...args) {
    const [top, left, bottom, right, vertical, horizontal, color, style] = args;
    this._sheet._borders.push({
      range: this.getA1Notation(), top, left, bottom, right, vertical, horizontal,
      color: color === undefined ? null : B.normalizeColor(color),
      style: style === undefined ? null : style,
    });
    return this;
  }

  // --- merges ---------------------------------------------------------------

  merge() {
    this._sheet._merge(this._rect());
    return this;
  }

  mergeAcross() {
    const rect = this._rect();
    for (let r = rect.r1; r <= rect.r2; r++) this._sheet._merge({r1: r, r2: r, c1: rect.c1, c2: rect.c2});
    return this;
  }

  mergeVertically() {
    const rect = this._rect();
    for (let c = rect.c1; c <= rect.c2; c++) this._sheet._merge({r1: rect.r1, r2: rect.r2, c1: c, c2: c});
    return this;
  }

  breakApart() {
    const rect = this._rect();
    const sheet = this._sheet;
    for (const m of sheet._merges) {
      if (intersects(m, rect) && !contains(rect, m)) {
        throw this._realm.error('You must select all cells in a merged range to merge or unmerge them.');
      }
    }
    sheet._merges = sheet._merges.filter((m) => !contains(rect, m));
    return this;
  }

  isPartOfMerge() {
    const rect = this._rect();
    return this._sheet._merges.some((m) => intersects(m, rect));
  }

  getMergedRanges() {
    const rect = this._rect();
    return this._realm.toContext(this._sheet._merges.filter((m) => intersects(m, rect))
        .map((m) => this._sheet.getRange(m.r1, m.c1, m.r2 - m.r1 + 1, m.c2 - m.c1 + 1)));
  }

  // --- misc -----------------------------------------------------------------

  protect() {
    const p = new B.ProtectionMock(this._sheet, 'RANGE', this);
    this._sheet._protections.push(p);
    return p;
  }

  createFilter() {
    if (this._sheet._filter) {
      throw this._realm.error('You can\'t create a filter in a sheet that already has a filter.');
    }
    this._sheet._filter = new FilterMock(this);
    return this._sheet._filter;
  }

  getFilter() {
    return this._sheet._filter;
  }

  applyRowBanding(theme) {
    const b = new BandingMock(this, theme || 'LIGHT_GREY', 'row');
    this._sheet._bandings.push(b);
    return b;
  }

  applyColumnBanding(theme) {
    const b = new BandingMock(this, theme || 'LIGHT_GREY', 'column');
    this._sheet._bandings.push(b);
    return b;
  }

  getBandings() {
    const rect = this._rect();
    return this._realm.toContext(this._sheet._bandings.filter((b) => intersects(b._range._rect(), rect)));
  }

  /**
   * sort(column) | sort({column, ascending}) | sort([specs...]); columns are
   * absolute sheet column numbers, as in Apps Script.
   * @param {*} sortSpecObj
   * @return {!RangeMock}
   */
  sort(sortSpecObj) {
    const specs = (Array.isArray(sortSpecObj) ? sortSpecObj : [sortSpecObj]).map((s) => (
      typeof s === 'number' ? {column: s, ascending: true} :
        {column: s.column, ascending: s.ascending !== false}));
    for (const s of specs) {
      if (s.column < this._col || s.column > this.getLastColumn()) {
        throw this._realm.error(`The column ${s.column} is out of range of the sort range.`);
      }
    }
    const sheet = this._sheet;
    const rows = [];
    for (let i = 0; i < this._numRows; i++) {
      const cells = [];
      for (let j = 0; j < this._numCols; j++) cells.push(sheet._cell(this._row + i, this._col + j));
      rows.push(cells);
    }
    rows.sort((a, b) => {
      for (const s of specs) {
        const idx = s.column - this._col;
        const cmp = compareCellValues(a[idx] ? a[idx].v : '', b[idx] ? b[idx].v : '', s.ascending);
        if (cmp !== 0) return cmp;
      }
      return 0;
    });
    rows.forEach((cells, i) => cells.forEach((cell, j) => sheet._setCell(this._row + i, this._col + j, cell)));
    sheet._touch();
    return this;
  }

  /**
   * copyTo(dest) | copyTo(dest, {contentsOnly, formatOnly}) |
   * copyTo(dest, copyPasteType, transposed)
   * @param {!RangeMock} destination
   * @param {*=} optionsOrType
   * @param {boolean=} transposed
   */
  copyTo(destination, optionsOrType, transposed) {
    let mode = 'all';
    if (typeof optionsOrType === 'string') {
      mode = {
        PASTE_NORMAL: 'all', PASTE_NO_BORDERS: 'all', PASTE_FORMAT: 'format',
        PASTE_FORMULA: 'formulas', PASTE_DATA_VALIDATION: 'validation', PASTE_VALUES: 'values',
        PASTE_CONDITIONAL_FORMATTING: 'none', PASTE_COLUMN_WIDTHS: 'widths',
      }[optionsOrType] || 'all';
    } else if (optionsOrType && typeof optionsOrType === 'object') {
      if (optionsOrType.contentsOnly) mode = 'formulas';
      else if (optionsOrType.formatOnly) mode = 'format';
    }
    this._sheet._paste(this, destination, mode, Boolean(transposed));
  }

  /**
   * Cuts and pastes (values, formats, notes, validation) to `target`.
   * @param {!RangeMock} target
   */
  moveTo(target) {
    const snapshot = this._map((cell) => cloneCell(cell));
    this._each((r, c) => this._sheet._setCell(r, c, undefined));
    this._sheet._touch();
    const dest = target._sheet;
    snapshot.forEach((row, i) => row.forEach((cell, j) => {
      dest._setCell(target._row + i, target._col + j, cell ? cloneCell(cell) : undefined);
    }));
    dest._touch();
  }

  /**
   * Deletes the cells, shifting the rest up (ROWS) or left (COLUMNS).
   * @param {string} shiftDimension SpreadsheetApp.Dimension value.
   */
  deleteCells(shiftDimension) {
    this._sheet._shiftCells(this._rect(), shiftDimension, -1);
  }

  /**
   * Inserts blank cells, shifting existing ones down (ROWS) or right (COLUMNS).
   * @param {string} shiftDimension
   * @return {!RangeMock}
   */
  insertCells(shiftDimension) {
    this._sheet._shiftCells(this._rect(), shiftDimension, 1);
    return this;
  }

  /** @return {string} */
  toString() {
    return 'Range';
  }
}

/** Format-related methods generated for RangeMock: [key, singular, plural, normalizer]. */
const FORMAT_METHODS = [
  ['background', 'Background', 'Backgrounds', B.normalizeColor],
  ['fontColor', 'FontColor', 'FontColors', B.normalizeColor],
  ['fontWeight', 'FontWeight', 'FontWeights', (v) => (v == null ? null : String(v).toLowerCase())],
  ['fontStyle', 'FontStyle', 'FontStyles', (v) => (v == null ? null : String(v).toLowerCase())],
  ['fontSize', 'FontSize', 'FontSizes', (v) => (v == null ? null : Number(v))],
  ['fontFamily', 'FontFamily', 'FontFamilies', (v) => (v == null ? null : String(v))],
  ['fontLine', 'FontLine', 'FontLines', (v) => (v == null ? null : String(v).toLowerCase())],
  ['horizontalAlignment', 'HorizontalAlignment', 'HorizontalAlignments',
    (v) => (v == null ? null : String(v).toLowerCase())],
  ['verticalAlignment', 'VerticalAlignment', 'VerticalAlignments',
    (v) => (v == null ? null : String(v).toLowerCase())],
  ['wrapStrategy', 'WrapStrategy', 'WrapStrategies', (v) => (v == null ? null : String(v))],
  ['numberFormat', 'NumberFormat', 'NumberFormats', (v) => (v == null ? null : String(v))],
  ['textDirection', 'TextDirection', 'TextDirections', (v) => (v == null ? null : String(v))],
];

for (const [key, singular, plural, normalize] of FORMAT_METHODS) {
  RangeMock.prototype[`set${singular}`] = function(value) {
    this._setFormat(key, normalize(value));
    return this;
  };
  RangeMock.prototype[`set${plural}`] = function(values) {
    this._check2D(values, `Range.set${plural}`);
    this._each((r, c, i, j) => this._sheet._setCellFormat(r, c, key, normalize(values[i][j])));
    return this;
  };
  RangeMock.prototype[`get${singular}`] = function() {
    return this._topLeft((cell) => RangeMock._formatOf(key, cell));
  };
  RangeMock.prototype[`get${plural}`] = function() {
    return this._realm.toContext(this._map((cell) => RangeMock._formatOf(key, cell)));
  };
}

// ---------------------------------------------------------------------------
// Sheet
// ---------------------------------------------------------------------------

/** Sheet mock. */
class SheetMock {
  /**
   * @param {!SpreadsheetMock} spreadsheet
   * @param {string} name
   * @param {number} sheetId
   * @param {!Object} env Shared runtime environment.
   * @param {{maxRows: (number|undefined), maxColumns: (number|undefined)}=} opts
   */
  constructor(spreadsheet, name, sheetId, env, opts = {}) {
    defineHidden(this, {
      _ss: spreadsheet, _env: env, _name: name, _sheetId: sheetId,
      _grid: [], _maxRows: opts.maxRows || DEFAULT_MAX_ROWS, _maxCols: opts.maxColumns || DEFAULT_MAX_COLS,
      _frozenRows: 0, _frozenCols: 0, _hiddenRows: new Set(), _hiddenCols: new Set(),
      _colWidths: new Map(), _rowHeights: new Map(), _forcedRowHeights: new Set(),
      _autoResizedCols: new Set(), _autoResizedRows: new Set(),
      _merges: [], _borders: [], _tabColor: null, _hidden: false, _filter: null,
      _cfRules: [], _protections: [], _bandings: [], _extents: null, _activeRange: null,
      _rightToLeft: false, _hiddenGridlines: false, _deleted: false, _developerMetadata: [],
    });
  }

  // --- internal storage -----------------------------------------------------

  /**
   * @param {number} r
   * @param {number} c
   * @return {?Object|undefined}
   */
  _cell(r, c) {
    const row = this._grid[r - 1];
    return row ? row[c - 1] : undefined;
  }

  /**
   * Returns the cell record, creating it (and growing the sheet) if needed.
   * @param {number} r
   * @param {number} c
   * @return {!Object}
   */
  _ensureCell(r, c) {
    let row = this._grid[r - 1];
    if (!row) {
      row = [];
      this._grid[r - 1] = row;
    }
    let cell = row[c - 1];
    if (!cell) {
      cell = newCell();
      row[c - 1] = cell;
    }
    if (r > this._maxRows) this._maxRows = r;
    if (c > this._maxCols) this._maxCols = c;
    return cell;
  }

  /**
   * Places a cell record (or removes it when undefined).
   * @param {number} r
   * @param {number} c
   * @param {?Object|undefined} cell
   */
  _setCell(r, c, cell) {
    if (!cell) {
      const row = this._grid[r - 1];
      if (row && row.length >= c) row[c - 1] = undefined;
      return;
    }
    if (!this._grid[r - 1]) this._grid[r - 1] = [];
    this._grid[r - 1][c - 1] = cell;
    if (r > this._maxRows) this._maxRows = r;
    if (c > this._maxCols) this._maxCols = c;
  }

  /** Invalidates cached extents after content changes. */
  _touch() {
    this._extents = null;
  }

  /** @return {{lastRow: number, lastCol: number}} */
  _computeExtents() {
    if (this._extents) return this._extents;
    let lastRow = 0;
    let lastCol = 0;
    for (let r = 0; r < this._grid.length; r++) {
      const row = this._grid[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        if (!isBlankCell(row[c])) {
          lastRow = r + 1;
          if (c + 1 > lastCol) lastCol = c + 1;
        }
      }
    }
    this._extents = {lastRow, lastCol};
    return this._extents;
  }

  /**
   * Converts a written value into a stored {v, f, link, numberFormat}.
   * @param {*} raw
   * @param {?Object} cell Existing cell (for its number format).
   * @param {number} r
   * @param {number} c
   * @return {{v: *, f: ?string, link: ?string, numberFormat: (string|undefined)}}
   */
  _coerce(raw, cell, r, c) {
    if (raw === null || raw === undefined) return {v: '', f: null, link: null};
    if (isDate(raw)) return {v: new Date(raw.getTime()), f: null, link: null};
    const type = typeof raw;
    if (type === 'number' || type === 'boolean') return {v: raw, f: null, link: null};
    if (type === 'bigint') return {v: Number(raw), f: null, link: null};
    if (type === 'string') {
      const plainText = cell && cell.fmt && cell.fmt.numberFormat === '@';
      if (plainText) return {v: raw, f: null, link: null};
      if (raw.length > 1 && raw[0] === '=') return this._coerceFormula(raw, r, c);
      if (raw[0] === '\'') return {v: raw.slice(1), f: null, link: null};
      const typed = autoType(raw);
      return {v: typed.value, f: null, link: null, numberFormat: typed.numberFormat};
    }
    const str = String(raw);
    this._env.warn('non-primitive-cell-value',
        `A ${typeName(raw)} value was written to '${this._name}'!${cellA1(r, c)}; Sheets stores ` +
        `its string form ${JSON.stringify(str.slice(0, 80))}.`, {sheet: this._name, cell: cellA1(r, c)});
    return {v: str, f: null, link: null};
  }

  /**
   * @param {*} formula
   * @param {number} r
   * @param {number} c
   * @return {{v: *, f: ?string, link: ?string}}
   */
  _coerceFormula(formula, r, c) {
    if (formula === null || formula === undefined || formula === '') return {v: '', f: null, link: null};
    const text = String(formula);
    if (text[0] !== '=') return this._coerce(text, this._cell(r, c), r, c);
    const result = evaluateFormula(text, this._env.options.formulaEvaluator || null,
        {sheet: this, row: r, column: c});
    if (result.unevaluated) {
      this._env.warn('formula-not-evaluated',
          `Formula ${JSON.stringify(text.slice(0, 80))} in '${this._name}'!${cellA1(r, c)} is not ` +
          'evaluated by the harness; its value reads as ' + JSON.stringify(result.value) + '.',
          {sheet: this._name, cell: cellA1(r, c), formula: text});
    }
    return {v: result.value, f: text, link: result.link || null};
  }

  /**
   * Writes a block of values atomically (validation errors write nothing).
   * @param {number} row
   * @param {number} col
   * @param {!Array<!Array<*>>} values
   * @param {string} mode 'value' | 'formula'
   */
  _writeBlock(row, col, values, mode) {
    const env = this._env;
    const tz = this._ss.getSpreadsheetTimeZone();
    const pending = [];
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const r = row + i;
        const c = col + j;
        const existing = this._cell(r, c);
        const co = mode === 'formula' ? this._coerceFormula(values[i][j], r, c) :
          this._coerce(values[i][j], existing, r, c);
        const dv = existing && existing.dv;
        if (dv && dv.getCriteriaType() === 'CHECKBOX' && !co.f && co.v === '') co.v = dv.uncheckedValue();
        if (dv && env.options.enforceDataValidation !== false && !dv.getAllowInvalid() && !co.f &&
            !dv.accepts(co.v, (v) => displayOf(v, existing.fmt, tz))) {
          throw env.realm.error(`The data you entered in cell ${cellA1(r, c)} violates the data ` +
              'validation rules set on this cell.');
        }
        pending.push([r, c, co]);
      }
    }
    for (const [r, c, co] of pending) {
      const cell = this._ensureCell(r, c);
      cell.v = co.v;
      cell.f = co.f || null;
      cell.link = co.link || null;
      cell.rich = null;
      if (co.numberFormat && !(cell.fmt && cell.fmt.numberFormat)) {
        cell.fmt = {...(cell.fmt || {}), numberFormat: co.numberFormat};
      }
    }
    this._touch();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {!Array<!Array<?B.RichTextValueMock>>} values
   */
  _writeRich(row, col, values) {
    values.forEach((rowValues, i) => rowValues.forEach((rich, j) => {
      const cell = this._ensureCell(row + i, col + j);
      if (rich === null || rich === undefined) {
        Object.assign(cell, {v: '', f: null, link: null, rich: null});
        return;
      }
      if (!(rich instanceof B.RichTextValueMock)) {
        throw this._env.realm.error(`The parameters (${typeName(rich)}) don't match the method ` +
            'signature for SpreadsheetApp.Range.setRichTextValue.');
      }
      Object.assign(cell, {v: rich.getText(), f: null, link: rich.getLinkUrl(), rich});
    }));
    this._touch();
  }

  /**
   * @param {?Object} cell
   * @return {?B.RichTextValueMock}
   */
  _richOf(cell) {
    if (cell && cell.rich) return cell.rich;
    const v = cell ? cell.v : '';
    if (typeof v !== 'string') return null;
    const links = cell && cell.link ? [{start: 0, end: v.length, url: cell.link}] : [];
    return new B.RichTextValueMock(this._env.realm, v, links, []);
  }

  /**
   * @param {number} r
   * @param {number} c
   * @param {string} key
   * @param {*} value null resets.
   */
  _setCellFormat(r, c, key, value) {
    if (value === null || value === undefined) {
      const cell = this._cell(r, c);
      if (cell && cell.fmt) {
        delete cell.fmt[key];
        if (Object.keys(cell.fmt).length === 0) cell.fmt = null;
      }
      return;
    }
    const cell = this._ensureCell(r, c);
    cell.fmt = cell.fmt || {};
    cell.fmt[key] = value;
  }

  /**
   * @param {number} r
   * @param {number} c
   * @param {?B.TextStyleMock} style
   */
  _applyTextStyle(r, c, style) {
    if (!style) {
      for (const key of ['fontWeight', 'fontStyle', 'fontLine', 'fontSize', 'fontFamily', 'fontColor']) {
        this._setCellFormat(r, c, key, null);
      }
      return;
    }
    const p = style.describe();
    if (p.bold !== undefined) this._setCellFormat(r, c, 'fontWeight', p.bold ? 'bold' : 'normal');
    if (p.italic !== undefined) this._setCellFormat(r, c, 'fontStyle', p.italic ? 'italic' : 'normal');
    if (p.underline !== undefined || p.strikethrough !== undefined) {
      this._setCellFormat(r, c, 'fontLine',
          p.strikethrough ? 'line-through' : (p.underline ? 'underline' : 'none'));
    }
    if (p.fontSize !== undefined) this._setCellFormat(r, c, 'fontSize', p.fontSize);
    if (p.fontFamily !== undefined) this._setCellFormat(r, c, 'fontFamily', p.fontFamily);
    if (p.foregroundColor !== undefined) this._setCellFormat(r, c, 'fontColor', p.foregroundColor);
  }

  /**
   * @param {?Object} cell
   * @return {!B.TextStyleMock}
   */
  _textStyleOf(cell) {
    const f = (key) => RangeMock._formatOf(key, cell);
    return new B.TextStyleMock({
      bold: f('fontWeight') === 'bold',
      italic: f('fontStyle') === 'italic',
      underline: f('fontLine') === 'underline',
      strikethrough: f('fontLine') === 'line-through',
      fontSize: f('fontSize'),
      fontFamily: f('fontFamily'),
      foregroundColor: f('fontColor'),
    });
  }

  /**
   * Applies (or clears) a validation rule on one cell, normalizing checkbox values.
   * @param {number} r
   * @param {number} c
   * @param {?B.DataValidationMock} rule
   */
  _applyValidation(r, c, rule) {
    if (!rule) {
      const cell = this._cell(r, c);
      if (cell) cell.dv = null;
      return;
    }
    const cell = this._ensureCell(r, c);
    cell.dv = rule;
    if (rule.isPlainCheckbox()) {
      if (cell.v === '' && !cell.f) cell.v = false;
    }
  }

  /**
   * @param {{r1: number, c1: number, r2: number, c2: number}} rect
   * @param {!Object=} options {contentsOnly, formatOnly, commentsOnly, validationsOnly}
   */
  _clearRect(rect, options) {
    const o = options || {};
    const any = o.contentsOnly || o.formatOnly || o.commentsOnly || o.validationsOnly;
    const contents = any ? Boolean(o.contentsOnly) : true;
    const formats = any ? Boolean(o.formatOnly) : true;
    const notes = Boolean(o.commentsOnly);
    const validations = Boolean(o.validationsOnly);
    for (let r = rect.r1; r <= rect.r2; r++) {
      const row = this._grid[r - 1];
      if (!row) continue;
      for (let c = rect.c1; c <= Math.min(rect.c2, row.length); c++) {
        const cell = row[c - 1];
        if (!cell) continue;
        if (validations) cell.dv = null;
        if (contents) {
          cell.v = cell.dv && cell.dv.getCriteriaType() === 'CHECKBOX' ? cell.dv.uncheckedValue() : '';
          cell.f = null;
          cell.link = null;
          cell.rich = null;
        }
        if (formats) cell.fmt = null;
        if (notes) cell.note = '';
      }
    }
    if (formats) {
      this._borders = this._borders.filter((b) => {
        const p = parseA1(b.range, this._maxRows, this._maxCols);
        return !p || !contains(rect, {r1: p.row, c1: p.col, r2: p.row + p.numRows - 1, c2: p.col + p.numCols - 1});
      });
    }
    this._touch();
  }

  /**
   * @param {{r1: number, c1: number, r2: number, c2: number}} rect
   */
  _merge(rect) {
    for (const m of this._merges) {
      if (intersects(m, rect) && !contains(rect, m)) {
        throw this._env.realm.error('You must select all cells in a merged range to merge or unmerge them.');
      }
    }
    this._merges = this._merges.filter((m) => !contains(rect, m));
    if (rect.r1 === rect.r2 && rect.c1 === rect.c2) return;
    this._merges.push({...rect});
    // Merging keeps only the top-left value.
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (r === rect.r1 && c === rect.c1) continue;
        const cell = this._cell(r, c);
        if (cell) Object.assign(cell, {v: '', f: null, link: null, rich: null});
      }
    }
    this._touch();
  }

  /**
   * Copies `src` into `dest` (possibly on another sheet).
   * @param {!RangeMock} src
   * @param {!RangeMock} dest
   * @param {string} mode all|values|formulas|format|validation|widths|none
   * @param {boolean} transposed
   */
  _paste(src, dest, mode, transposed) {
    let snapshot = src._map((cell) => cloneCell(cell));
    if (transposed) snapshot = snapshot[0].map((_, j) => snapshot.map((row) => row[j]));
    const srcRows = snapshot.length;
    const srcCols = snapshot[0].length;
    const tile = (destN, srcN) => (destN > srcN && destN % srcN === 0 ? destN : srcN);
    const rows = tile(dest._numRows, srcRows);
    const cols = tile(dest._numCols, srcCols);
    const target = dest._sheet;
    if (mode === 'widths') {
      for (let j = 0; j < cols; j++) {
        target.setColumnWidth(dest._col + j, src._sheet.getColumnWidth(src._col + (j % srcCols)));
      }
      return;
    }
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const from = snapshot[i % srcRows][j % srcCols];
        const r = dest._row + i;
        const c = dest._col + j;
        if (mode === 'all') {
          target._setCell(r, c, from ? cloneCell(from) : undefined);
          continue;
        }
        if (!from && mode !== 'values' && mode !== 'formulas') continue;
        const cell = target._ensureCell(r, c);
        if (mode === 'values') {
          Object.assign(cell, {v: from ? from.v : '', f: null, link: from ? from.link : null, rich: null});
        } else if (mode === 'formulas') {
          Object.assign(cell, {v: from ? from.v : '', f: from ? from.f : null,
            link: from ? from.link : null, rich: from ? from.rich : null});
        } else if (mode === 'format') {
          cell.fmt = from.fmt ? {...from.fmt} : null;
          cell.dv = from.dv;
        } else if (mode === 'validation') {
          cell.dv = from.dv;
        }
      }
    }
    target._touch();
  }

  /**
   * Shifts cells for Range.deleteCells/insertCells.
   * @param {{r1: number, c1: number, r2: number, c2: number}} rect
   * @param {string} dimension 'ROWS' | 'COLUMNS'
   * @param {number} direction -1 delete, +1 insert
   */
  _shiftCells(rect, dimension, direction) {
    if (dimension === 'ROWS') {
      const n = rect.r2 - rect.r1 + 1;
      const last = Math.max(this._grid.length, rect.r2) + n;
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (direction < 0) {
          for (let r = rect.r1; r <= last; r++) this._setCell(r, c, this._cell(r + n, c));
        } else {
          for (let r = last; r >= rect.r1 + n; r--) this._setCell(r, c, this._cell(r - n, c));
          for (let r = rect.r1; r < rect.r1 + n; r++) this._setCell(r, c, undefined);
        }
      }
    } else {
      const n = rect.c2 - rect.c1 + 1;
      for (let r = rect.r1; r <= rect.r2; r++) {
        const row = this._grid[r - 1];
        if (!row) continue;
        const last = Math.max(row.length, rect.c2) + n;
        if (direction < 0) {
          for (let c = rect.c1; c <= last; c++) this._setCell(r, c, this._cell(r, c + n));
        } else {
          for (let c = last; c >= rect.c1 + n; c--) this._setCell(r, c, this._cell(r, c - n));
          for (let c = rect.c1; c < rect.c1 + n; c++) this._setCell(r, c, undefined);
        }
      }
    }
    this._touch();
  }

  /**
   * Creates a range, validating arguments like Sheet.getRange.
   * @param {number} row
   * @param {number} col
   * @param {number} numRows
   * @param {number} numCols
   * @param {?string=} hint
   * @return {!RangeMock}
   */
  _makeRange(row, col, numRows, numCols, hint = null) {
    const realm = this._env.realm;
    if (row < 1) throw realm.error('The starting row of the range is too small.');
    if (col < 1) throw realm.error('The starting column of the range is too small.');
    if (numRows < 1) throw realm.error('The number of rows in the range must be at least 1.');
    if (numCols < 1) throw realm.error('The number of columns in the range must be at least 1.');
    if (this._env.options.strictBounds &&
        (row + numRows - 1 > this._maxRows || col + numCols - 1 > this._maxCols)) {
      throw realm.error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    return new RangeMock(this, row, col, numRows, numCols, hint);
  }

  /**
   * Inserts rows before `at`, inheriting formats/validation from `inheritRow`.
   * @param {number} at
   * @param {number} n
   * @param {number} inheritRow
   */
  _insertRows(at, n, inheritRow) {
    const realm = this._env.realm;
    if (!(n >= 1)) throw realm.error('The number of rows to insert must be at least 1.');
    if (at < 1 || at > this._maxRows + 1) throw realm.error('Those rows are out of bounds.');
    const templateRow = inheritRow >= 1 && inheritRow <= this._maxRows ? (this._grid[inheritRow - 1] || []) : [];
    const template = templateRow.map((cell) => (cell && (cell.fmt || cell.dv) ?
      {fmt: cell.fmt ? {...cell.fmt} : null, dv: cell.dv} : undefined));
    insertHoles(this._grid, at - 1, n);
    this._maxRows += n;
    template.forEach((t, j) => {
      if (!t) return;
      for (let r = at; r < at + n; r++) {
        const cell = this._ensureCell(r, j + 1);
        cell.fmt = t.fmt ? {...t.fmt} : null;
        cell.dv = t.dv;
        if (t.dv && t.dv.getCriteriaType() === 'CHECKBOX') cell.v = t.dv.uncheckedValue();
      }
    });
    this._hiddenRows = shiftKeys(this._hiddenRows, at, n);
    this._rowHeights = shiftKeys(this._rowHeights, at, n);
    this._forcedRowHeights = shiftKeys(this._forcedRowHeights, at, n);
    for (const m of this._merges) {
      if (m.r1 >= at) {
        m.r1 += n;
        m.r2 += n;
      } else if (m.r2 >= at) {
        m.r2 += n;
      }
    }
    if (at <= this._frozenRows) this._frozenRows += n;
    this._touch();
  }

  /**
   * @param {number} at
   * @param {number} n
   */
  _deleteRows(at, n) {
    const realm = this._env.realm;
    if (at < 1 || !(n >= 1) || at + n - 1 > this._maxRows) throw realm.error('Those rows are out of bounds.');
    if (this._maxRows - n <= this._frozenRows) {
      throw realm.error('Sorry, it is not possible to delete all non-frozen rows.');
    }
    this._grid.splice(at - 1, n);
    this._maxRows -= n;
    this._hiddenRows = shiftKeys(this._hiddenRows, at, -n);
    this._rowHeights = shiftKeys(this._rowHeights, at, -n);
    this._forcedRowHeights = shiftKeys(this._forcedRowHeights, at, -n);
    const end = at + n - 1;
    this._merges = this._merges.filter((m) => !(m.r1 >= at && m.r2 <= end)).map((m) => {
      const overlap = Math.max(0, Math.min(m.r2, end) - Math.max(m.r1, at) + 1);
      if (m.r1 > end) return {...m, r1: m.r1 - n, r2: m.r2 - n};
      if (overlap > 0) return {...m, r1: Math.min(m.r1, at), r2: m.r2 - overlap};
      return m;
    }).filter((m) => m.r2 > m.r1 || m.c2 > m.c1);
    if (at <= this._frozenRows) this._frozenRows -= Math.min(n, this._frozenRows - at + 1);
    this._touch();
  }

  /**
   * Inserts columns before `at`, inheriting formats/validation from `inheritCol`.
   * @param {number} at
   * @param {number} n
   * @param {number} inheritCol
   */
  _insertCols(at, n, inheritCol) {
    const realm = this._env.realm;
    if (!(n >= 1)) throw realm.error('The number of columns to insert must be at least 1.');
    if (at < 1 || at > this._maxCols + 1) throw realm.error('Those columns are out of bounds.');
    const template = [];
    if (inheritCol >= 1 && inheritCol <= this._maxCols) {
      this._grid.forEach((row, i) => {
        const cell = row && row[inheritCol - 1];
        if (cell && (cell.fmt || cell.dv)) template.push({r: i + 1, fmt: cell.fmt, dv: cell.dv});
      });
    }
    for (const row of this._grid) if (row) insertHoles(row, at - 1, n);
    this._maxCols += n;
    for (const t of template) {
      for (let c = at; c < at + n; c++) {
        const cell = this._ensureCell(t.r, c);
        cell.fmt = t.fmt ? {...t.fmt} : null;
        cell.dv = t.dv;
        if (t.dv && t.dv.getCriteriaType() === 'CHECKBOX') cell.v = t.dv.uncheckedValue();
      }
    }
    this._hiddenCols = shiftKeys(this._hiddenCols, at, n);
    this._colWidths = shiftKeys(this._colWidths, at, n);
    this._autoResizedCols = shiftKeys(this._autoResizedCols, at, n);
    for (const m of this._merges) {
      if (m.c1 >= at) {
        m.c1 += n;
        m.c2 += n;
      } else if (m.c2 >= at) {
        m.c2 += n;
      }
    }
    if (at <= this._frozenCols) this._frozenCols += n;
    this._touch();
  }

  /**
   * @param {number} at
   * @param {number} n
   */
  _deleteCols(at, n) {
    const realm = this._env.realm;
    if (at < 1 || !(n >= 1) || at + n - 1 > this._maxCols) {
      throw realm.error('Those columns are out of bounds.');
    }
    if (this._maxCols - n <= this._frozenCols) {
      throw realm.error('Sorry, it is not possible to delete all non-frozen columns.');
    }
    for (const row of this._grid) if (row) row.splice(at - 1, n);
    this._maxCols -= n;
    this._hiddenCols = shiftKeys(this._hiddenCols, at, -n);
    this._colWidths = shiftKeys(this._colWidths, at, -n);
    this._autoResizedCols = shiftKeys(this._autoResizedCols, at, -n);
    const end = at + n - 1;
    this._merges = this._merges.filter((m) => !(m.c1 >= at && m.c2 <= end)).map((m) => {
      const overlap = Math.max(0, Math.min(m.c2, end) - Math.max(m.c1, at) + 1);
      if (m.c1 > end) return {...m, c1: m.c1 - n, c2: m.c2 - n};
      if (overlap > 0) return {...m, c1: Math.min(m.c1, at), c2: m.c2 - overlap};
      return m;
    }).filter((m) => m.r2 > m.r1 || m.c2 > m.c1);
    if (at <= this._frozenCols) this._frozenCols -= Math.min(n, this._frozenCols - at + 1);
    this._touch();
  }

  /**
   * @param {number} start
   * @param {number} n
   * @param {boolean} isRow
   */
  _checkLines(start, n, isRow) {
    const max = isRow ? this._maxRows : this._maxCols;
    if (!(start >= 1) || !(n >= 1) || start + n - 1 > max) {
      throw this._env.realm.error(isRow ? 'Those rows are out of bounds.' : 'Those columns are out of bounds.');
    }
  }

  // --- identity -------------------------------------------------------------

  getName() {
    return this._name;
  }

  getSheetName() {
    return this._name;
  }

  /**
   * @param {string} name
   * @return {!SheetMock}
   */
  setName(name) {
    const newName = String(name);
    if (newName === this._name) return this;
    if (this._ss._sheets.some((s) => s !== this && s._name.toLowerCase() === newName.toLowerCase())) {
      throw this._env.realm.error(
          `A sheet with the name "${newName}" already exists. Please enter another name.`);
    }
    this._name = newName;
    return this;
  }

  getSheetId() {
    return this._sheetId;
  }

  getIndex() {
    return this._ss._sheets.indexOf(this) + 1;
  }

  getParent() {
    return this._ss;
  }

  getType() {
    return 'GRID';
  }

  getFormUrl() {
    return null;
  }

  // --- dimensions -----------------------------------------------------------

  getLastRow() {
    return this._computeExtents().lastRow;
  }

  getLastColumn() {
    return this._computeExtents().lastCol;
  }

  getMaxRows() {
    return this._maxRows;
  }

  getMaxColumns() {
    return this._maxCols;
  }

  /**
   * getRange(a1) | getRange(row, column[, numRows[, numColumns]])
   * @param {...*} args
   * @return {!RangeMock}
   */
  getRange(...args) {
    const realm = this._env.realm;
    if (typeof args[0] === 'string' && args.length === 1) {
      const parsed = parseA1(args[0], this._maxRows, this._maxCols);
      if (!parsed || (parsed.sheetName !== null && parsed.sheetName !== this._name)) {
        throw realm.error('Range not found');
      }
      return this._makeRange(parsed.row, parsed.col, parsed.numRows, parsed.numCols, parsed.hint);
    }
    const [row, col, numRows = 1, numCols = 1] = args;
    const nums = [row, col, numRows, numCols];
    if (args.length < 2 || args.length > 4 || nums.some((n) => typeof n !== 'number' || Number.isNaN(n))) {
      throw realm.error(`The parameters (${args.map(typeName).join(',')}) don't match the method ` +
          'signature for SpreadsheetApp.Sheet.getRange.');
    }
    return this._makeRange(Math.trunc(row), Math.trunc(col), Math.trunc(numRows), Math.trunc(numCols));
  }

  /**
   * @param {!Array<string>} a1Notations
   * @return {!Object} Minimal RangeList.
   */
  getRangeList(a1Notations) {
    const ranges = [...a1Notations].map((a1) => this.getRange(a1));
    const realm = this._env.realm;
    const list = {
      getRanges: () => realm.toContext([...ranges]),
      clear: (o) => (ranges.forEach((r) => r.clear(o)), list),
      clearContent: () => (ranges.forEach((r) => r.clearContent()), list),
      clearFormat: () => (ranges.forEach((r) => r.clearFormat()), list),
      setValue: (v) => (ranges.forEach((r) => r.setValue(v)), list),
      setBackground: (v) => (ranges.forEach((r) => r.setBackground(v)), list),
      setFontWeight: (v) => (ranges.forEach((r) => r.setFontWeight(v)), list),
      setNumberFormat: (v) => (ranges.forEach((r) => r.setNumberFormat(v)), list),
      insertCheckboxes: () => (ranges.forEach((r) => r.insertCheckboxes()), list),
      removeCheckboxes: () => (ranges.forEach((r) => r.removeCheckboxes()), list),
      check: () => (ranges.forEach((r) => r.check()), list),
      uncheck: () => (ranges.forEach((r) => r.uncheck()), list),
      activate: () => list,
      toString: () => 'RangeList',
    };
    return list;
  }

  getDataRange() {
    const {lastRow, lastCol} = this._computeExtents();
    return this._makeRange(1, 1, Math.max(1, lastRow), Math.max(1, lastCol));
  }

  /**
   * @param {number} startRow
   * @param {number} startColumn
   * @param {number} numRows
   * @param {number} numColumns
   * @return {!Array<!Array<*>>}
   */
  getSheetValues(startRow, startColumn, numRows, numColumns) {
    const nr = numRows === -1 ? Math.max(1, this.getLastRow() - startRow + 1) : numRows;
    const nc = numColumns === -1 ? Math.max(1, this.getLastColumn() - startColumn + 1) : numColumns;
    return this.getRange(startRow, startColumn, nr, nc).getValues();
  }

  /**
   * Appends a row after the last row with content (growing the sheet).
   * @param {!Array<*>} rowContents
   * @return {!SheetMock}
   */
  appendRow(rowContents) {
    if (!Array.isArray(rowContents)) {
      throw this._env.realm.error(`The parameters (${typeName(rowContents)}) don't match the method ` +
          'signature for SpreadsheetApp.Sheet.appendRow.');
    }
    if (rowContents.length === 0) return this;
    this._writeBlock(this.getLastRow() + 1, 1, [rowContents], 'value');
    return this;
  }

  insertRows(rowIndex, numRows = 1) {
    this._insertRows(rowIndex, numRows, rowIndex);
  }

  insertRowBefore(beforePosition) {
    this._insertRows(beforePosition, 1, beforePosition);
    return this;
  }

  insertRowAfter(afterPosition) {
    this._insertRows(afterPosition + 1, 1, afterPosition);
    return this;
  }

  insertRowsBefore(beforePosition, howMany) {
    this._insertRows(beforePosition, howMany, beforePosition);
    return this;
  }

  insertRowsAfter(afterPosition, howMany) {
    this._insertRows(afterPosition + 1, howMany, afterPosition);
    return this;
  }

  insertColumns(columnIndex, numColumns = 1) {
    this._insertCols(columnIndex, numColumns, columnIndex);
  }

  insertColumnBefore(beforePosition) {
    this._insertCols(beforePosition, 1, beforePosition);
    return this;
  }

  insertColumnAfter(afterPosition) {
    this._insertCols(afterPosition + 1, 1, afterPosition);
    return this;
  }

  insertColumnsBefore(beforePosition, howMany) {
    this._insertCols(beforePosition, howMany, beforePosition);
    return this;
  }

  insertColumnsAfter(afterPosition, howMany) {
    this._insertCols(afterPosition + 1, howMany, afterPosition);
    return this;
  }

  deleteRow(rowPosition) {
    this._deleteRows(rowPosition, 1);
    return this;
  }

  deleteRows(rowPosition, howMany) {
    this._deleteRows(rowPosition, howMany);
  }

  deleteColumn(columnPosition) {
    this._deleteCols(columnPosition, 1);
    return this;
  }

  deleteColumns(columnPosition, howMany) {
    this._deleteCols(columnPosition, howMany);
  }

  // --- clearing -------------------------------------------------------------

  /**
   * @param {!Object=} options {contentsOnly, formatOnly, commentsOnly, validationsOnly}
   * @return {!SheetMock}
   */
  clear(options) {
    let widest = this._maxCols;
    for (const row of this._grid) if (row && row.length > widest) widest = row.length;
    this._clearRect({r1: 1, c1: 1, r2: Math.max(this._maxRows, this._grid.length), c2: widest}, options);
    return this;
  }

  clearContents() {
    return this.clear({contentsOnly: true});
  }

  clearFormats() {
    return this.clear({formatOnly: true});
  }

  clearNotes() {
    return this.clear({commentsOnly: true});
  }

  clearConditionalFormatRules() {
    this._cfRules = [];
  }

  // --- freezing, tabs, visibility ----------------------------------------------

  setFrozenRows(rows) {
    if (!(rows >= 0)) throw this._env.realm.error('The number of frozen rows must be at least 0.');
    this._frozenRows = rows;
  }

  getFrozenRows() {
    return this._frozenRows;
  }

  setFrozenColumns(columns) {
    if (!(columns >= 0)) throw this._env.realm.error('The number of frozen columns must be at least 0.');
    this._frozenCols = columns;
  }

  getFrozenColumns() {
    return this._frozenCols;
  }

  setTabColor(color) {
    this._tabColor = B.normalizeColor(color);
    return this;
  }

  getTabColor() {
    return this._tabColor;
  }

  hideColumns(columnIndex, numColumns = 1) {
    this._checkLines(columnIndex, numColumns, false);
    for (let c = columnIndex; c < columnIndex + numColumns; c++) this._hiddenCols.add(c);
  }

  showColumns(columnIndex, numColumns = 1) {
    this._checkLines(columnIndex, numColumns, false);
    for (let c = columnIndex; c < columnIndex + numColumns; c++) this._hiddenCols.delete(c);
  }

  hideColumn(column) {
    this.hideColumns(column.getColumn(), column.getNumColumns());
  }

  unhideColumn(column) {
    this.showColumns(column.getColumn(), column.getNumColumns());
  }

  isColumnHiddenByUser(columnPosition) {
    return this._hiddenCols.has(columnPosition);
  }

  hideRows(rowIndex, numRows = 1) {
    this._checkLines(rowIndex, numRows, true);
    for (let r = rowIndex; r < rowIndex + numRows; r++) this._hiddenRows.add(r);
  }

  showRows(rowIndex, numRows = 1) {
    this._checkLines(rowIndex, numRows, true);
    for (let r = rowIndex; r < rowIndex + numRows; r++) this._hiddenRows.delete(r);
  }

  hideRow(row) {
    this.hideRows(row.getRow(), row.getNumRows());
  }

  unhideRow(row) {
    this.showRows(row.getRow(), row.getNumRows());
  }

  isRowHiddenByUser(rowPosition) {
    return this._hiddenRows.has(rowPosition);
  }

  isRowHiddenByFilter() {
    return false;
  }

  hideSheet() {
    const visible = this._ss._sheets.filter((s) => !s._hidden && s !== this);
    if (visible.length === 0) throw this._env.realm.error('You can\'t hide all the sheets in a document.');
    this._hidden = true;
    return this;
  }

  showSheet() {
    this._hidden = false;
    return this;
  }

  isSheetHidden() {
    return this._hidden;
  }

  activate() {
    this._ss._activeSheet = this;
    this._hidden = false;
    return this;
  }

  // --- sizes ----------------------------------------------------------------

  setColumnWidth(columnPosition, width) {
    this._checkLines(columnPosition, 1, false);
    this._colWidths.set(columnPosition, width);
    return this;
  }

  setColumnWidths(startColumn, numColumns, width) {
    this._checkLines(startColumn, numColumns, false);
    for (let c = startColumn; c < startColumn + numColumns; c++) this._colWidths.set(c, width);
    return this;
  }

  getColumnWidth(columnPosition) {
    return this._colWidths.has(columnPosition) ? this._colWidths.get(columnPosition) : DEFAULT_COLUMN_WIDTH;
  }

  setRowHeight(rowPosition, height) {
    this._checkLines(rowPosition, 1, true);
    this._rowHeights.set(rowPosition, height);
    return this;
  }

  setRowHeights(startRow, numRows, height) {
    this._checkLines(startRow, numRows, true);
    for (let r = startRow; r < startRow + numRows; r++) this._rowHeights.set(r, height);
    return this;
  }

  setRowHeightsForced(startRow, numRows, height) {
    this._checkLines(startRow, numRows, true);
    for (let r = startRow; r < startRow + numRows; r++) {
      this._rowHeights.set(r, height);
      this._forcedRowHeights.add(r);
    }
    return this;
  }

  getRowHeight(rowPosition) {
    return this._rowHeights.has(rowPosition) ? this._rowHeights.get(rowPosition) : DEFAULT_ROW_HEIGHT;
  }

  autoResizeColumn(columnPosition) {
    return this.autoResizeColumns(columnPosition, 1);
  }

  /**
   * Approximates content width (7px per character + padding).
   * @param {number} startColumn
   * @param {number} numColumns
   * @return {!SheetMock}
   */
  autoResizeColumns(startColumn, numColumns) {
    this._checkLines(startColumn, numColumns, false);
    const tz = this._ss.getSpreadsheetTimeZone();
    for (let c = startColumn; c < startColumn + numColumns; c++) {
      let maxLen = 0;
      for (const row of this._grid) {
        const cell = row && row[c - 1];
        if (cell) maxLen = Math.max(maxLen, ...displayOf(cell.v, cell.fmt, tz).split('\n').map((s) => s.length));
      }
      this._colWidths.set(c, Math.max(30, Math.round(maxLen * 7 + 10)));
      this._autoResizedCols.add(c);
    }
    return this;
  }

  autoResizeRows(startRow, numRows) {
    this._checkLines(startRow, numRows, true);
    for (let r = startRow; r < startRow + numRows; r++) {
      this._autoResizedRows.add(r);
      this._rowHeights.delete(r);
    }
    return this;
  }

  // --- selection --------------------------------------------------------------

  getActiveRange() {
    return this._activeRange || this.getRange(1, 1);
  }

  getActiveCell() {
    const r = this.getActiveRange();
    return this.getRange(r.getRow(), r.getColumn());
  }

  getCurrentCell() {
    return this.getActiveCell();
  }

  setActiveRange(range) {
    this._activeRange = range;
    this._ss._activeSheet = this;
    return range;
  }

  setActiveSelection(range) {
    return this.setActiveRange(typeof range === 'string' ? this.getRange(range) : range);
  }

  setCurrentCell(cell) {
    return this.setActiveRange(cell);
  }

  getSelection() {
    return this._ss.getSelection();
  }

  // --- filters, protections, conditional formats ---------------------------------

  getFilter() {
    return this._filter;
  }

  protect() {
    const p = new B.ProtectionMock(this, 'SHEET', null);
    this._protections.push(p);
    return p;
  }

  getProtections(type) {
    return this._env.realm.toContext(this._protections.filter((p) => !type || p.getProtectionType() === type));
  }

  getConditionalFormatRules() {
    return this._env.realm.toContext([...this._cfRules]);
  }

  setConditionalFormatRules(rules) {
    this._cfRules = [...rules];
  }

  getBandings() {
    return this._env.realm.toContext([...this._bandings]);
  }

  getCharts() {
    return this._env.realm.toContext([]);
  }

  getImages() {
    return this._env.realm.toContext([]);
  }

  getNamedRanges() {
    return this._env.realm.toContext([...this._ss._namedRanges.values()]
        .filter((nr) => nr.getRange().getSheet() === this));
  }

  getDeveloperMetadata() {
    return this._env.realm.toContext([...this._developerMetadata]);
  }

  addDeveloperMetadata(key, value) {
    this._developerMetadata.push({getKey: () => key, getValue: () => value ?? null, toString: () => 'DeveloperMetadata'});
    return this;
  }

  setRightToLeft(rightToLeft) {
    this._rightToLeft = Boolean(rightToLeft);
    return this;
  }

  isRightToLeft() {
    return this._rightToLeft;
  }

  setHiddenGridlines(hide) {
    this._hiddenGridlines = Boolean(hide);
    return this;
  }

  hasHiddenGridlines() {
    return this._hiddenGridlines;
  }

  /**
   * Sorts rows below the frozen rows by a column.
   * @param {number} columnPosition
   * @param {boolean=} ascending
   * @return {!SheetMock}
   */
  sort(columnPosition, ascending = true) {
    const {lastRow, lastCol} = this._computeExtents();
    const first = this._frozenRows + 1;
    if (lastRow >= first) {
      this.getRange(first, 1, lastRow - first + 1, Math.max(lastCol, columnPosition))
          .sort({column: columnPosition, ascending});
    }
    return this;
  }

  /**
   * Copies this sheet into another (or the same) spreadsheet.
   * @param {!SpreadsheetMock} spreadsheet
   * @return {!SheetMock}
   */
  copyTo(spreadsheet) {
    let name = `Copy of ${this._name}`;
    let n = 2;
    while (spreadsheet._sheets.some((s) => s._name.toLowerCase() === name.toLowerCase())) {
      name = `Copy of ${this._name} ${n++}`;
    }
    const copy = spreadsheet._createSheet(name, spreadsheet._sheets.length, false);
    copy._copyFrom(this);
    return copy;
  }

  /**
   * Copies contents, formats and layout from another sheet.
   * @param {!SheetMock} other
   */
  _copyFrom(other) {
    this._grid = other._grid.map((row) => (row ? row.map((cell) => cloneCell(cell)) : row));
    this._maxRows = other._maxRows;
    this._maxCols = other._maxCols;
    this._frozenRows = other._frozenRows;
    this._frozenCols = other._frozenCols;
    this._hiddenRows = new Set(other._hiddenRows);
    this._hiddenCols = new Set(other._hiddenCols);
    this._colWidths = new Map(other._colWidths);
    this._rowHeights = new Map(other._rowHeights);
    this._merges = other._merges.map((m) => ({...m}));
    this._tabColor = other._tabColor;
    this._touch();
  }

  /** @return {string} */
  toString() {
    return 'Sheet';
  }
}

// ---------------------------------------------------------------------------
// Spreadsheet
// ---------------------------------------------------------------------------

/** Spreadsheet mock. */
class SpreadsheetMock {
  /**
   * @param {!Object} env
   * @param {{id: string, name: string, timeZone: string, locale: (string|undefined)}} props
   */
  constructor(env, props) {
    defineHidden(this, {
      _env: env, _id: props.id, _name: props.name, _timeZone: props.timeZone,
      _locale: props.locale || 'en_US', _sheets: [], _activeSheet: null, _nextSheetId: 0,
      _namedRanges: new Map(), _recalc: 'ON_CHANGE',
    });
  }

  /**
   * @param {string} name
   * @param {number} index 0-based insert position.
   * @param {boolean} activate
   * @param {!Object=} opts
   * @return {!SheetMock}
   */
  _createSheet(name, index, activate, opts = {}) {
    const id = opts.sheetId !== undefined ? opts.sheetId :
      (this._nextSheetId === 0 ? 0 : 1000000 + this._nextSheetId);
    this._nextSheetId++;
    const sheet = new SheetMock(this, name, id, this._env, opts);
    const pos = Math.max(0, Math.min(index, this._sheets.length));
    this._sheets.splice(pos, 0, sheet);
    if (activate || !this._activeSheet) this._activeSheet = sheet;
    return sheet;
  }

  getId() {
    return this._id;
  }

  getName() {
    return this._name;
  }

  rename(newName) {
    this._name = String(newName);
  }

  getUrl() {
    return `https://docs.google.com/spreadsheets/d/${this._id}/edit`;
  }

  getSheets() {
    return this._env.realm.toContext([...this._sheets]);
  }

  getNumSheets() {
    return this._sheets.length;
  }

  /**
   * @param {string} name
   * @return {?SheetMock}
   */
  getSheetByName(name) {
    return this._sheets.find((s) => s._name === String(name)) || null;
  }

  getSheetById(id) {
    return this._sheets.find((s) => s._sheetId === id) || null;
  }

  /**
   * insertSheet() | (index) | (name) | (name, index) | (options) | (index, options) |
   * (name, options) | (name, index, options)
   * @param {...*} args
   * @return {!SheetMock}
   */
  insertSheet(...args) {
    let name = null;
    let index = null;
    let options = null;
    for (const a of args) {
      if (typeof a === 'string') name = a;
      else if (typeof a === 'number') index = a;
      else if (a && typeof a === 'object') options = a;
    }
    if (name === null) {
      let n = this._sheets.length + 1;
      while (this._sheets.some((s) => s._name.toLowerCase() === `sheet${n}`)) n++;
      name = `Sheet${n}`;
    }
    if (this._sheets.some((s) => s._name.toLowerCase() === name.toLowerCase())) {
      throw this._env.realm.error(`A sheet with the name "${name}" already exists. Please enter another name.`);
    }
    const sheet = this._createSheet(name, index === null ? this._sheets.length : index, true);
    if (options && options.template) sheet._copyFrom(options.template);
    return sheet;
  }

  /** @param {!SheetMock} sheet */
  deleteSheet(sheet) {
    if (!(sheet instanceof SheetMock) || !this._sheets.includes(sheet)) {
      throw this._env.realm.error('The parameters (' + typeName(sheet) +
          ') don\'t match the method signature for SpreadsheetApp.Spreadsheet.deleteSheet.');
    }
    if (this._sheets.length === 1) {
      throw this._env.realm.error('You can\'t remove all the sheets in a document.');
    }
    this._sheets = this._sheets.filter((s) => s !== sheet);
    sheet._deleted = true;
    if (this._activeSheet === sheet) this._activeSheet = this._sheets[0];
    for (const [name, nr] of this._namedRanges) {
      if (nr.getRange().getSheet() === sheet) this._namedRanges.delete(name);
    }
  }

  duplicateActiveSheet() {
    const copy = this.getActiveSheet().copyTo(this);
    this._sheets = this._sheets.filter((s) => s !== copy);
    this._sheets.splice(this.getActiveSheet().getIndex(), 0, copy);
    this._activeSheet = copy;
    return copy;
  }

  moveActiveSheet(pos) {
    const sheet = this.getActiveSheet();
    this._sheets = this._sheets.filter((s) => s !== sheet);
    this._sheets.splice(Math.max(0, pos - 1), 0, sheet);
  }

  renameActiveSheet(newName) {
    this.getActiveSheet().setName(newName);
  }

  getActiveSheet() {
    return this._activeSheet || this._sheets[0] || null;
  }

  setActiveSheet(sheet) {
    this._activeSheet = sheet;
    return sheet;
  }

  getActiveRange() {
    const sheet = this.getActiveSheet();
    return sheet ? sheet.getActiveRange() : null;
  }

  getActiveCell() {
    const sheet = this.getActiveSheet();
    return sheet ? sheet.getActiveCell() : null;
  }

  getCurrentCell() {
    return this.getActiveCell();
  }

  setActiveRange(range) {
    return range.getSheet().setActiveRange(range);
  }

  setActiveSelection(range) {
    return this.getActiveSheet().setActiveSelection(range);
  }

  setCurrentCell(cell) {
    return cell.getSheet().setActiveRange(cell);
  }

  getSelection() {
    const ss = this;
    return {
      getActiveSheet: () => ss.getActiveSheet(),
      getActiveRange: () => ss.getActiveRange(),
      getCurrentCell: () => ss.getActiveCell(),
      getActiveRangeList: () => null,
      toString: () => 'Selection',
    };
  }

  /**
   * Records a toast (see runtime.ui.toasts).
   * @param {string} msg
   * @param {string=} title
   * @param {number=} timeoutSeconds
   */
  toast(msg, title, timeoutSeconds) {
    const record = {
      type: 'toast',
      message: String(msg),
      title: title === undefined || title === null ? '' : String(title),
      timeoutSeconds: timeoutSeconds === undefined ? 5 : timeoutSeconds,
    };
    this._env.ui.toasts.push(record);
    this._env.ui.history.push(record);
  }

  getSpreadsheetTimeZone() {
    return this._timeZone;
  }

  setSpreadsheetTimeZone(timezone) {
    this._timeZone = String(timezone);
  }

  getSpreadsheetLocale() {
    return this._locale;
  }

  setSpreadsheetLocale(locale) {
    this._locale = String(locale);
  }

  getRecalculationInterval() {
    return this._recalc;
  }

  setRecalculationInterval(interval) {
    this._recalc = interval;
    return this;
  }

  /**
   * @param {string} a1Notation Optionally sheet-qualified.
   * @return {!RangeMock}
   */
  getRange(a1Notation) {
    const realm = this._env.realm;
    const sheetGuess = this.getActiveSheet();
    const parsed = parseA1(a1Notation, sheetGuess ? sheetGuess._maxRows : DEFAULT_MAX_ROWS,
        sheetGuess ? sheetGuess._maxCols : DEFAULT_MAX_COLS);
    if (!parsed) throw realm.error('Range not found');
    const sheet = parsed.sheetName === null ? sheetGuess : this.getSheetByName(parsed.sheetName);
    if (!sheet) throw realm.error('Range not found');
    const reparsed = parseA1(a1Notation, sheet._maxRows, sheet._maxCols);
    return sheet._makeRange(reparsed.row, reparsed.col, reparsed.numRows, reparsed.numCols, reparsed.hint);
  }

  getRangeByName(name) {
    const nr = this._namedRanges.get(String(name));
    return nr ? nr.getRange() : null;
  }

  setNamedRange(name, range) {
    this._namedRanges.set(String(name), new B.NamedRangeMock(this, String(name), range));
  }

  getNamedRanges() {
    return this._env.realm.toContext([...this._namedRanges.values()]);
  }

  removeNamedRange(name) {
    this._namedRanges.delete(String(name));
  }

  getLastRow() {
    return this.getActiveSheet().getLastRow();
  }

  getLastColumn() {
    return this.getActiveSheet().getLastColumn();
  }

  getDataRange() {
    return this.getActiveSheet().getDataRange();
  }

  appendRow(rowContents) {
    return this.getActiveSheet().appendRow(rowContents);
  }

  getOwner() {
    const email = this._env.userEmail;
    return {getEmail: () => email, toString: () => 'User'};
  }

  getEditors() {
    return this._env.realm.toContext([]);
  }

  getViewers() {
    return this._env.realm.toContext([]);
  }

  addEditor() {
    return this;
  }

  addViewer() {
    return this;
  }

  getFormUrl() {
    return null;
  }

  getDeveloperMetadata() {
    return this._env.realm.toContext([]);
  }

  toString() {
    return 'Spreadsheet';
  }
}

// ---------------------------------------------------------------------------
// SpreadsheetApp
// ---------------------------------------------------------------------------

/** The global `SpreadsheetApp`. */
class SpreadsheetAppMock {
  /**
   * @param {!Object} env
   * @param {!SpreadsheetMock} active
   */
  constructor(env, active) {
    defineHidden(this, {_env: env, _active: active, _flushes: 0});
    Object.assign(this, B.ENUMS);
  }

  getActive() {
    return this._active;
  }

  getActiveSpreadsheet() {
    return this._active;
  }

  setActiveSpreadsheet(spreadsheet) {
    this._active = spreadsheet;
  }

  getActiveSheet() {
    return this._active.getActiveSheet();
  }

  setActiveSheet(sheet) {
    return this._active.setActiveSheet(sheet);
  }

  getActiveRange() {
    return this._active.getActiveRange();
  }

  setActiveRange(range) {
    return this._active.setActiveRange(range);
  }

  getCurrentCell() {
    return this._active.getCurrentCell();
  }

  setCurrentCell(cell) {
    return this._active.setCurrentCell(cell);
  }

  getSelection() {
    return this._active.getSelection();
  }

  getUi() {
    return this._env.ui.instance;
  }

  newDataValidation() {
    return new B.DataValidationBuilderMock(this._env.realm);
  }

  newRichTextValue() {
    return new B.RichTextValueBuilderMock(this._env.realm);
  }

  newTextStyle() {
    return new B.TextStyleBuilderMock();
  }

  newConditionalFormatRule() {
    return new B.ConditionalFormatRuleBuilderMock(this._env.realm);
  }

  newFilterCriteria() {
    return new B.FilterCriteriaBuilderMock(this._env.realm);
  }

  newColor() {
    let hex = '#000000';
    const builder = {
      setRgbColor: (c) => ((hex = B.normalizeColor(c)), builder),
      build: () => ({asRgbColor: () => ({asHexString: () => hex}), asHexString: () => hex,
        getColorType: () => 'RGB', toString: () => 'Color'}),
      toString: () => 'ColorBuilder',
    };
    return builder;
  }

  /** Records a flush (no-op in memory). */
  flush() {
    this._flushes++;
  }

  /**
   * @param {string} id
   * @return {!SpreadsheetMock}
   */
  openById(id) {
    const ss = this._env.spreadsheets.get(String(id));
    if (!ss) {
      throw this._env.realm.error(
          'Unexpected error while getting the method or property openById on object SpreadsheetApp.');
    }
    return ss;
  }

  /**
   * @param {string} url
   * @return {!SpreadsheetMock}
   */
  openByUrl(url) {
    const m = /\/d\/([^/]+)/.exec(String(url));
    const ss = m && this._env.spreadsheets.get(m[1]);
    if (!ss) {
      throw this._env.realm.error(
          'Unexpected error while getting the method or property openByUrl on object SpreadsheetApp.');
    }
    return ss;
  }

  /**
   * Creates and registers a new in-memory spreadsheet.
   * @param {string} name
   * @param {number=} rows
   * @param {number=} columns
   * @return {!SpreadsheetMock}
   */
  create(name, rows, columns) {
    const id = `created-${this._env.spreadsheets.size + 1}-${Date.now().toString(36)}`;
    const ss = new SpreadsheetMock(this._env, {id, name: String(name), timeZone: this._active._timeZone});
    ss._createSheet('Sheet1', 0, true, {maxRows: rows, maxColumns: columns});
    this._env.spreadsheets.set(id, ss);
    return ss;
  }

  enableAllDataSourcesExecution() {}

  enableBigQueryExecution() {}

  toString() {
    return 'SpreadsheetApp';
  }
}

module.exports = {
  SpreadsheetAppMock,
  SpreadsheetMock,
  SheetMock,
  RangeMock,
  FilterMock,
  FORMAT_DEFAULTS,
  DEFAULT_NUMBER_FORMAT,
  autoType,
  evaluateFormula,
  displayOf,
  isBlankCell,
  cloneCell,
};
