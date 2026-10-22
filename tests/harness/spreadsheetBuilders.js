/**
 * @fileoverview Enums, builders and small value classes used by the
 * SpreadsheetApp mock: DataValidation(Builder), RichTextValue(Builder),
 * TextStyle(Builder), ConditionalFormatRule(Builder), FilterCriteria(Builder),
 * Protection and NamedRange.
 */
'use strict';

const {defineHidden} = require('./uiMock');

/**
 * Builds a frozen enum whose values are their own names (Apps Script enums
 * stringify to their names, so `==` comparisons and logging behave alike).
 * @param {...string} names
 * @return {!Object<string, string>}
 */
function enumOf(...names) {
  const out = {};
  for (const name of names) out[name] = name;
  return Object.freeze(out);
}

const ENUMS = Object.freeze({
  BorderStyle: enumOf('DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE'),
  WrapStrategy: enumOf('WRAP', 'OVERFLOW', 'CLIP'),
  DataValidationCriteria: enumOf(
      'DATE_AFTER', 'DATE_BEFORE', 'DATE_BETWEEN', 'DATE_EQUAL_TO', 'DATE_IS_VALID_DATE',
      'DATE_NOT_BETWEEN', 'DATE_ON_OR_AFTER', 'DATE_ON_OR_BEFORE', 'NUMBER_BETWEEN',
      'NUMBER_EQUAL_TO', 'NUMBER_GREATER_THAN', 'NUMBER_GREATER_THAN_OR_EQUAL_TO',
      'NUMBER_LESS_THAN', 'NUMBER_LESS_THAN_OR_EQUAL_TO', 'NUMBER_NOT_BETWEEN',
      'NUMBER_NOT_EQUAL_TO', 'TEXT_CONTAINS', 'TEXT_DOES_NOT_CONTAIN', 'TEXT_EQUAL_TO',
      'TEXT_IS_VALID_EMAIL', 'TEXT_IS_VALID_URL', 'VALUE_IN_LIST', 'VALUE_IN_RANGE',
      'CUSTOM_FORMULA', 'CHECKBOX'),
  Dimension: enumOf('COLUMNS', 'ROWS'),
  CopyPasteType: enumOf(
      'PASTE_NORMAL', 'PASTE_NO_BORDERS', 'PASTE_FORMAT', 'PASTE_FORMULA',
      'PASTE_DATA_VALIDATION', 'PASTE_VALUES', 'PASTE_CONDITIONAL_FORMATTING',
      'PASTE_COLUMN_WIDTHS'),
  BandingTheme: enumOf(
      'LIGHT_GREY', 'CYAN', 'GREEN', 'YELLOW', 'ORANGE', 'BLUE', 'TEAL', 'GREY', 'BROWN',
      'LIGHT_GREEN', 'INDIGO', 'PINK'),
  ProtectionType: enumOf('RANGE', 'SHEET'),
  SheetType: enumOf('GRID', 'OBJECT', 'DATASOURCE'),
  TextDirection: enumOf('LEFT_TO_RIGHT', 'RIGHT_TO_LEFT'),
  Direction: enumOf('UP', 'DOWN', 'PREVIOUS', 'NEXT'),
  AutoFillSeries: enumOf('DEFAULT_SERIES', 'ALTERNATE_SERIES'),
  DeveloperMetadataVisibility: enumOf('DOCUMENT', 'PROJECT'),
  RecalculationInterval: enumOf('ON_CHANGE', 'MINUTE', 'HOUR'),
  BooleanCriteria: enumOf(
      'CELL_EMPTY', 'CELL_NOT_EMPTY', 'DATE_AFTER', 'DATE_BEFORE', 'DATE_EQUAL_TO',
      'DATE_AFTER_RELATIVE', 'DATE_BEFORE_RELATIVE', 'DATE_EQUAL_TO_RELATIVE', 'NUMBER_BETWEEN',
      'NUMBER_EQUAL_TO', 'NUMBER_GREATER_THAN', 'NUMBER_GREATER_THAN_OR_EQUAL_TO',
      'NUMBER_LESS_THAN', 'NUMBER_LESS_THAN_OR_EQUAL_TO', 'NUMBER_NOT_BETWEEN',
      'NUMBER_NOT_EQUAL_TO', 'TEXT_CONTAINS', 'TEXT_DOES_NOT_CONTAIN', 'TEXT_EQUAL_TO',
      'TEXT_NOT_EQUAL_TO', 'TEXT_STARTS_WITH', 'TEXT_ENDS_WITH', 'CUSTOM_FORMULA'),
  ColorType: enumOf('UNSUPPORTED', 'RGB', 'THEME'),
  GroupControlTogglePosition: enumOf('BEFORE', 'AFTER'),
});

const NAMED_COLORS = Object.freeze({
  white: '#ffffff', black: '#000000', red: '#ff0000', green: '#008000', blue: '#0000ff',
  yellow: '#ffff00', gray: '#808080', grey: '#808080', orange: '#ffa500', purple: '#800080',
  pink: '#ffc0cb', cyan: '#00ffff', aqua: '#00ffff', magenta: '#ff00ff', fuchsia: '#ff00ff',
  lime: '#00ff00', silver: '#c0c0c0', navy: '#000080', teal: '#008080', maroon: '#800000',
  olive: '#808000', brown: '#a52a2a', lightgray: '#d3d3d3', lightgrey: '#d3d3d3',
  darkgray: '#a9a9a9', darkgrey: '#a9a9a9', lightgreen: '#90ee90', lightblue: '#add8e6',
});

/**
 * Normalizes a CSS color like Sheets does when reading it back
 * ('#ABC' -> '#aabbcc', 'red' -> '#ff0000', 'rgb(0,0,255)' -> '#0000ff').
 * @param {*} color
 * @return {?string} null for null/empty (meaning "reset to default").
 */
function normalizeColor(color) {
  if (color === null || color === undefined || color === '') return null;
  const s = String(color).trim().toLowerCase();
  if (NAMED_COLORS[s]) return NAMED_COLORS[s];
  if (/^#[0-9a-f]{3}$/.test(s)) return `#${s[1]}${s[1]}${s[2]}${s[2]}${s[3]}${s[3]}`;
  if (/^#[0-9a-f]{6}$/.test(s)) return s;
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (rgb) {
    return '#' + [rgb[1], rgb[2], rgb[3]]
        .map((v) => Math.max(0, Math.min(255, Number(v))).toString(16).padStart(2, '0')).join('');
  }
  return s;
}

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------

/**
 * Immutable data validation rule (SpreadsheetApp DataValidation).
 */
class DataValidationMock {
  /**
   * @param {!Object} realm
   * @param {{criteria: string, values: !Array, allowInvalid: boolean, helpText: string}} spec
   */
  constructor(realm, spec) {
    defineHidden(this, {
      _realm: realm,
      _criteria: spec.criteria,
      _values: [...spec.values],
      _allowInvalid: spec.allowInvalid,
      _helpText: spec.helpText,
    });
  }

  /** @return {string} A DataValidationCriteria value. */
  getCriteriaType() {
    return this._criteria;
  }

  /** @return {!Array} Criteria arguments (e.g. [list, showDropdown]). */
  getCriteriaValues() {
    return this._realm.toContext(this._values.map((v) => (Array.isArray(v) ? [...v] : v)));
  }

  /** @return {boolean} */
  getAllowInvalid() {
    return this._allowInvalid;
  }

  /** @return {string} */
  getHelpText() {
    return this._helpText;
  }

  /** @return {!DataValidationBuilderMock} A builder pre-populated with this rule. */
  copy() {
    const builder = new DataValidationBuilderMock(this._realm);
    builder._criteria = this._criteria;
    builder._values = [...this._values];
    builder._allowInvalid = this._allowInvalid;
    builder._helpText = this._helpText;
    return builder;
  }

  /** @return {string} */
  toString() {
    return 'DataValidation';
  }

  /**
   * Host-side description used by dumps/fixtures.
   * @return {!Object}
   */
  describe() {
    const values = this._values.map((v) => {
      if (v && typeof v.getA1Notation === 'function') {
        return {$range: `${v.getSheet().getName()}!${v.getA1Notation()}`};
      }
      if (Array.isArray(v)) return [...v];
      return v;
    });
    return {
      criteria: this._criteria,
      values,
      allowInvalid: this._allowInvalid,
      helpText: this._helpText,
    };
  }

  /** @return {boolean} True for checkbox rules without custom values. */
  isPlainCheckbox() {
    return this._criteria === 'CHECKBOX' && this._values.length === 0;
  }

  /** @return {*} The value meaning "checked". */
  checkedValue() {
    return this._values.length > 0 ? this._values[0] : true;
  }

  /** @return {*} The value meaning "unchecked". */
  uncheckedValue() {
    if (this._values.length === 0) return false;
    return this._values.length > 1 ? this._values[1] : '';
  }

  /**
   * Tests whether a stored value satisfies the rule. Only criteria that can be
   * evaluated without a formula engine are checked; everything else passes.
   * @param {*} value
   * @param {function(*): string} display Formats a value as displayed text.
   * @return {boolean}
   */
  accepts(value, display) {
    if (value === '' || value === null || value === undefined) return true;
    const vals = this._values;
    switch (this._criteria) {
      case 'CHECKBOX':
        if (this._values.length === 0) return value === true || value === false;
        return String(value) === String(this.checkedValue()) ||
            String(value) === String(this.uncheckedValue());
      case 'VALUE_IN_LIST':
        return (vals[0] || []).map(String).includes(display(value));
      case 'VALUE_IN_RANGE': {
        const range = vals[0];
        if (!range || typeof range.getDisplayValues !== 'function') return true;
        const allowed = this._realm.toHost(range.getDisplayValues()).flat().map(String);
        return allowed.includes(display(value));
      }
      case 'NUMBER_BETWEEN':
        return typeof value === 'number' && value >= vals[0] && value <= vals[1];
      case 'NUMBER_NOT_BETWEEN':
        return typeof value === 'number' && (value < vals[0] || value > vals[1]);
      case 'NUMBER_EQUAL_TO': return typeof value === 'number' && value === vals[0];
      case 'NUMBER_NOT_EQUAL_TO': return typeof value === 'number' && value !== vals[0];
      case 'NUMBER_GREATER_THAN': return typeof value === 'number' && value > vals[0];
      case 'NUMBER_GREATER_THAN_OR_EQUAL_TO': return typeof value === 'number' && value >= vals[0];
      case 'NUMBER_LESS_THAN': return typeof value === 'number' && value < vals[0];
      case 'NUMBER_LESS_THAN_OR_EQUAL_TO': return typeof value === 'number' && value <= vals[0];
      case 'TEXT_CONTAINS': return String(value).includes(String(vals[0]));
      case 'TEXT_DOES_NOT_CONTAIN': return !String(value).includes(String(vals[0]));
      case 'TEXT_EQUAL_TO': return String(value) === String(vals[0]);
      case 'TEXT_IS_VALID_EMAIL': return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value));
      case 'TEXT_IS_VALID_URL': return /^(https?|ftp):\/\/\S+$/i.test(String(value));
      default:
        return true;
    }
  }
}

/**
 * DataValidationBuilder mock (SpreadsheetApp.newDataValidation()).
 */
class DataValidationBuilderMock {
  /** @param {!Object} realm */
  constructor(realm) {
    defineHidden(this, {
      _realm: realm, _criteria: null, _values: [], _allowInvalid: true, _helpText: '',
    });
  }

  /**
   * @param {string} criteria
   * @param {!Array} values
   * @return {!DataValidationBuilderMock}
   */
  _set(criteria, values) {
    this._criteria = criteria;
    this._values = values;
    return this;
  }

  /**
   * @param {*=} checkedValue
   * @param {*=} uncheckedValue
   * @return {!DataValidationBuilderMock}
   */
  requireCheckbox(...args) {
    return this._set('CHECKBOX', args.slice(0, 2));
  }

  /**
   * @param {!Array<string>} values
   * @param {boolean=} showDropdown
   * @return {!DataValidationBuilderMock}
   */
  requireValueInList(values, showDropdown = true) {
    if (!Array.isArray(values)) {
      throw this._realm.error(`The parameters (${typeof values}) don't match the method ` +
          'signature for SpreadsheetApp.DataValidationBuilder.requireValueInList.');
    }
    return this._set('VALUE_IN_LIST', [this._realm.toHost(values).slice(), Boolean(showDropdown)]);
  }

  /**
   * @param {!Object} range A RangeMock.
   * @param {boolean=} showDropdown
   * @return {!DataValidationBuilderMock}
   */
  requireValueInRange(range, showDropdown = true) {
    if (!range || typeof range.getA1Notation !== 'function') {
      throw this._realm.error(`The parameters (${typeof range}) don't match the method ` +
          'signature for SpreadsheetApp.DataValidationBuilder.requireValueInRange.');
    }
    return this._set('VALUE_IN_RANGE', [range, Boolean(showDropdown)]);
  }

  requireNumberBetween(start, end) {
    return this._set('NUMBER_BETWEEN', [start, end]);
  }

  requireNumberNotBetween(start, end) {
    return this._set('NUMBER_NOT_BETWEEN', [start, end]);
  }

  requireNumberEqualTo(n) {
    return this._set('NUMBER_EQUAL_TO', [n]);
  }

  requireNumberNotEqualTo(n) {
    return this._set('NUMBER_NOT_EQUAL_TO', [n]);
  }

  requireNumberGreaterThan(n) {
    return this._set('NUMBER_GREATER_THAN', [n]);
  }

  requireNumberGreaterThanOrEqualTo(n) {
    return this._set('NUMBER_GREATER_THAN_OR_EQUAL_TO', [n]);
  }

  requireNumberLessThan(n) {
    return this._set('NUMBER_LESS_THAN', [n]);
  }

  requireNumberLessThanOrEqualTo(n) {
    return this._set('NUMBER_LESS_THAN_OR_EQUAL_TO', [n]);
  }

  requireTextContains(text) {
    return this._set('TEXT_CONTAINS', [text]);
  }

  requireTextDoesNotContain(text) {
    return this._set('TEXT_DOES_NOT_CONTAIN', [text]);
  }

  requireTextEqualTo(text) {
    return this._set('TEXT_EQUAL_TO', [text]);
  }

  requireTextIsEmail() {
    return this._set('TEXT_IS_VALID_EMAIL', []);
  }

  requireTextIsUrl() {
    return this._set('TEXT_IS_VALID_URL', []);
  }

  requireDate() {
    return this._set('DATE_IS_VALID_DATE', []);
  }

  requireDateAfter(date) {
    return this._set('DATE_AFTER', [date]);
  }

  requireDateBefore(date) {
    return this._set('DATE_BEFORE', [date]);
  }

  requireDateBetween(start, end) {
    return this._set('DATE_BETWEEN', [start, end]);
  }

  requireDateNotBetween(start, end) {
    return this._set('DATE_NOT_BETWEEN', [start, end]);
  }

  requireDateEqualTo(date) {
    return this._set('DATE_EQUAL_TO', [date]);
  }

  requireDateOnOrAfter(date) {
    return this._set('DATE_ON_OR_AFTER', [date]);
  }

  requireDateOnOrBefore(date) {
    return this._set('DATE_ON_OR_BEFORE', [date]);
  }

  requireFormulaSatisfied(formula) {
    return this._set('CUSTOM_FORMULA', [String(formula)]);
  }

  /**
   * @param {string} criteria A DataValidationCriteria value.
   * @param {!Array} args
   * @return {!DataValidationBuilderMock}
   */
  withCriteria(criteria, args) {
    return this._set(String(criteria), this._realm.toHost(args || []).slice());
  }

  /**
   * @param {boolean} allowInvalidData
   * @return {!DataValidationBuilderMock}
   */
  setAllowInvalid(allowInvalidData) {
    this._allowInvalid = Boolean(allowInvalidData);
    return this;
  }

  /**
   * @param {string} helpText
   * @return {!DataValidationBuilderMock}
   */
  setHelpText(helpText) {
    this._helpText = helpText === null || helpText === undefined ? '' : String(helpText);
    return this;
  }

  getCriteriaType() {
    return this._criteria;
  }

  getCriteriaValues() {
    return this._realm.toContext([...this._values]);
  }

  getAllowInvalid() {
    return this._allowInvalid;
  }

  getHelpText() {
    return this._helpText;
  }

  /** @return {!DataValidationBuilderMock} */
  copy() {
    const builder = new DataValidationBuilderMock(this._realm);
    builder._criteria = this._criteria;
    builder._values = [...this._values];
    builder._allowInvalid = this._allowInvalid;
    builder._helpText = this._helpText;
    return builder;
  }

  /** @return {!DataValidationMock} */
  build() {
    if (!this._criteria) {
      throw this._realm.error('You must specify criteria for the data validation rule.');
    }
    return new DataValidationMock(this._realm, {
      criteria: this._criteria, values: this._values,
      allowInvalid: this._allowInvalid, helpText: this._helpText,
    });
  }

  /** @return {string} */
  toString() {
    return 'DataValidationBuilder';
  }
}

// ---------------------------------------------------------------------------
// Text styles and rich text
// ---------------------------------------------------------------------------

/** Immutable TextStyle. */
class TextStyleMock {
  /**
   * @param {!Object} props
   */
  constructor(props) {
    defineHidden(this, {_props: {...props}});
  }

  getFontFamily() {
    return this._props.fontFamily ?? null;
  }

  getFontSize() {
    return this._props.fontSize ?? null;
  }

  getForegroundColor() {
    return this._props.foregroundColor ?? null;
  }

  isBold() {
    return this._props.bold ?? null;
  }

  isItalic() {
    return this._props.italic ?? null;
  }

  isStrikethrough() {
    return this._props.strikethrough ?? null;
  }

  isUnderline() {
    return this._props.underline ?? null;
  }

  /** @return {!TextStyleBuilderMock} */
  copy() {
    const b = new TextStyleBuilderMock();
    Object.assign(b._props, this._props);
    return b;
  }

  /** @return {!Object} */
  describe() {
    return {...this._props};
  }

  toString() {
    return 'TextStyle';
  }
}

/** TextStyleBuilder (SpreadsheetApp.newTextStyle()). */
class TextStyleBuilderMock {
  constructor() {
    defineHidden(this, {_props: {}});
  }

  setBold(v) {
    this._props.bold = Boolean(v);
    return this;
  }

  setItalic(v) {
    this._props.italic = Boolean(v);
    return this;
  }

  setUnderline(v) {
    this._props.underline = Boolean(v);
    return this;
  }

  setStrikethrough(v) {
    this._props.strikethrough = Boolean(v);
    return this;
  }

  setFontSize(v) {
    this._props.fontSize = Number(v);
    return this;
  }

  setFontFamily(v) {
    this._props.fontFamily = String(v);
    return this;
  }

  setForegroundColor(v) {
    this._props.foregroundColor = normalizeColor(v);
    return this;
  }

  build() {
    return new TextStyleMock(this._props);
  }

  toString() {
    return 'TextStyleBuilder';
  }
}

/** Immutable RichTextValue. */
class RichTextValueMock {
  /**
   * @param {!Object} realm
   * @param {string} text
   * @param {!Array<{start: number, end: number, url: string}>} links
   * @param {!Array<{start: number, end: number, style: !TextStyleMock}>} styles
   * @param {number=} offset Start index of this run within the parent value.
   */
  constructor(realm, text, links, styles, offset = 0) {
    defineHidden(this, {_realm: realm, _text: text, _links: links, _styles: styles, _offset: offset});
  }

  /** @return {string} */
  getText() {
    return this._text;
  }

  /**
   * getLinkUrl() returns the link covering the whole text (else null);
   * getLinkUrl(start, end) returns the link covering that span.
   * @param {number=} startOffset
   * @param {number=} endOffset
   * @return {?string}
   */
  getLinkUrl(startOffset, endOffset) {
    const start = startOffset === undefined ? 0 : startOffset;
    const end = endOffset === undefined ? this._text.length : endOffset;
    const link = this._links.find((l) => l.start <= start && l.end >= end);
    return link ? link.url : null;
  }

  /** @return {number} */
  getStartIndex() {
    return this._offset;
  }

  /** @return {number} */
  getEndIndex() {
    return this._offset + this._text.length;
  }

  /**
   * @param {number=} startOffset
   * @param {number=} endOffset
   * @return {!TextStyleMock}
   */
  getTextStyle(startOffset, endOffset) {
    const start = startOffset === undefined ? 0 : startOffset;
    const end = endOffset === undefined ? this._text.length : endOffset;
    const style = this._styles.find((s) => s.start <= start && s.end >= end);
    return style ? style.style : new TextStyleMock({});
  }

  /** @return {!Array<!RichTextValueMock>} Runs split at link/style boundaries. */
  getRuns() {
    const cuts = new Set([0, this._text.length]);
    for (const l of [...this._links, ...this._styles]) {
      cuts.add(l.start);
      cuts.add(l.end);
    }
    const points = [...cuts].filter((p) => p >= 0 && p <= this._text.length).sort((a, b) => a - b);
    const runs = [];
    for (let i = 0; i < points.length - 1; i++) {
      const [s, e] = [points[i], points[i + 1]];
      if (s === e) continue;
      const links = this._links.filter((l) => l.start <= s && l.end >= e)
          .map((l) => ({start: 0, end: e - s, url: l.url}));
      const styles = this._styles.filter((st) => st.start <= s && st.end >= e)
          .map((st) => ({start: 0, end: e - s, style: st.style}));
      runs.push(new RichTextValueMock(this._realm, this._text.slice(s, e), links, styles, s));
    }
    return this._realm.toContext(runs);
  }

  /** @return {!RichTextValueBuilderMock} */
  copy() {
    const b = new RichTextValueBuilderMock(this._realm);
    b._text = this._text;
    b._links = this._links.map((l) => ({...l}));
    b._styles = this._styles.map((s) => ({...s}));
    return b;
  }

  /** @return {!Object} Host-side description for dumps. */
  describe() {
    return {text: this._text, links: this._links.map((l) => ({...l}))};
  }

  toString() {
    return 'RichTextValue';
  }
}

/** RichTextValueBuilder (SpreadsheetApp.newRichTextValue()). */
class RichTextValueBuilderMock {
  /** @param {!Object} realm */
  constructor(realm) {
    defineHidden(this, {_realm: realm, _text: '', _links: [], _styles: []});
  }

  /**
   * @param {string} text
   * @return {!RichTextValueBuilderMock}
   */
  setText(text) {
    this._text = String(text);
    return this;
  }

  /**
   * setLinkUrl(url) | setLinkUrl(startOffset, endOffset, url)
   * @param {...*} args
   * @return {!RichTextValueBuilderMock}
   */
  setLinkUrl(...args) {
    if (args.length >= 3) {
      const [start, end, url] = args;
      this._links = this._links.filter((l) => l.end <= start || l.start >= end);
      if (url !== null && url !== undefined) this._links.push({start, end, url: String(url)});
    } else {
      const [url] = args;
      this._links = url === null || url === undefined ? [] :
          [{start: 0, end: this._text.length, url: String(url)}];
    }
    return this;
  }

  /**
   * setTextStyle(style) | setTextStyle(startOffset, endOffset, style)
   * @param {...*} args
   * @return {!RichTextValueBuilderMock}
   */
  setTextStyle(...args) {
    if (args.length >= 3) {
      const [start, end, style] = args;
      this._styles.push({start, end, style});
    } else {
      this._styles = [{start: 0, end: this._text.length, style: args[0]}];
    }
    return this;
  }

  /** @return {!RichTextValueMock} */
  build() {
    // Whole-text links set before setText() should span the final text.
    const links = this._links.map((l) => (l.start === 0 && l.end === 0 ?
      {...l, end: this._text.length} : {...l}));
    return new RichTextValueMock(this._realm, this._text, links, this._styles.map((s) => ({...s})));
  }

  toString() {
    return 'RichTextValueBuilder';
  }
}

// ---------------------------------------------------------------------------
// Conditional formatting and filters (recorded, not evaluated)
// ---------------------------------------------------------------------------

/** ConditionalFormatRule (immutable). */
class ConditionalFormatRuleMock {
  /**
   * @param {!Object} realm
   * @param {!Object} spec
   */
  constructor(realm, spec) {
    defineHidden(this, {_realm: realm, _spec: {...spec, ranges: [...spec.ranges]}});
  }

  getRanges() {
    return this._realm.toContext([...this._spec.ranges]);
  }

  getBooleanCondition() {
    const spec = this._spec;
    const realm = this._realm;
    return {
      getCriteriaType: () => spec.criteria,
      getCriteriaValues: () => realm.toContext([...spec.values]),
      getBackground: () => spec.format.background ?? null,
      getFontColor: () => spec.format.fontColor ?? null,
      getBold: () => spec.format.bold ?? null,
      getItalic: () => spec.format.italic ?? null,
      getStrikethrough: () => spec.format.strikethrough ?? null,
      getUnderline: () => spec.format.underline ?? null,
      toString: () => 'BooleanCondition',
    };
  }

  getGradientCondition() {
    return null;
  }

  copy() {
    const b = new ConditionalFormatRuleBuilderMock(this._realm);
    b._spec = {...this._spec, ranges: [...this._spec.ranges], format: {...this._spec.format}};
    return b;
  }

  /** @return {!Object} */
  describe() {
    return {
      criteria: this._spec.criteria,
      values: [...this._spec.values],
      format: {...this._spec.format},
      ranges: this._spec.ranges.map((r) => r.getA1Notation()),
    };
  }

  toString() {
    return 'ConditionalFormatRule';
  }
}

/** ConditionalFormatRuleBuilder (SpreadsheetApp.newConditionalFormatRule()). */
class ConditionalFormatRuleBuilderMock {
  /** @param {!Object} realm */
  constructor(realm) {
    defineHidden(this, {
      _realm: realm, _spec: {criteria: null, values: [], format: {}, ranges: []},
    });
  }

  _when(criteria, values) {
    this._spec.criteria = criteria;
    this._spec.values = values;
    return this;
  }

  whenCellEmpty() {
    return this._when('CELL_EMPTY', []);
  }

  whenCellNotEmpty() {
    return this._when('CELL_NOT_EMPTY', []);
  }

  whenTextEqualTo(t) {
    return this._when('TEXT_EQUAL_TO', [t]);
  }

  whenTextNotEqualTo(t) {
    return this._when('TEXT_NOT_EQUAL_TO', [t]);
  }

  whenTextContains(t) {
    return this._when('TEXT_CONTAINS', [t]);
  }

  whenTextDoesNotContain(t) {
    return this._when('TEXT_DOES_NOT_CONTAIN', [t]);
  }

  whenTextStartsWith(t) {
    return this._when('TEXT_STARTS_WITH', [t]);
  }

  whenTextEndsWith(t) {
    return this._when('TEXT_ENDS_WITH', [t]);
  }

  whenFormulaSatisfied(f) {
    return this._when('CUSTOM_FORMULA', [f]);
  }

  whenNumberBetween(a, b) {
    return this._when('NUMBER_BETWEEN', [a, b]);
  }

  whenNumberNotBetween(a, b) {
    return this._when('NUMBER_NOT_BETWEEN', [a, b]);
  }

  whenNumberEqualTo(n) {
    return this._when('NUMBER_EQUAL_TO', [n]);
  }

  whenNumberNotEqualTo(n) {
    return this._when('NUMBER_NOT_EQUAL_TO', [n]);
  }

  whenNumberGreaterThan(n) {
    return this._when('NUMBER_GREATER_THAN', [n]);
  }

  whenNumberGreaterThanOrEqualTo(n) {
    return this._when('NUMBER_GREATER_THAN_OR_EQUAL_TO', [n]);
  }

  whenNumberLessThan(n) {
    return this._when('NUMBER_LESS_THAN', [n]);
  }

  whenNumberLessThanOrEqualTo(n) {
    return this._when('NUMBER_LESS_THAN_OR_EQUAL_TO', [n]);
  }

  whenDateAfter(d) {
    return this._when('DATE_AFTER', [d]);
  }

  whenDateBefore(d) {
    return this._when('DATE_BEFORE', [d]);
  }

  whenDateEqualTo(d) {
    return this._when('DATE_EQUAL_TO', [d]);
  }

  setBackground(c) {
    this._spec.format.background = normalizeColor(c);
    return this;
  }

  setFontColor(c) {
    this._spec.format.fontColor = normalizeColor(c);
    return this;
  }

  setBold(b) {
    this._spec.format.bold = Boolean(b);
    return this;
  }

  setItalic(b) {
    this._spec.format.italic = Boolean(b);
    return this;
  }

  setStrikethrough(b) {
    this._spec.format.strikethrough = Boolean(b);
    return this;
  }

  setUnderline(b) {
    this._spec.format.underline = Boolean(b);
    return this;
  }

  setRanges(ranges) {
    this._spec.ranges = [...ranges];
    return this;
  }

  getRanges() {
    return this._realm.toContext([...this._spec.ranges]);
  }

  copy() {
    const b = new ConditionalFormatRuleBuilderMock(this._realm);
    b._spec = {...this._spec, ranges: [...this._spec.ranges], format: {...this._spec.format}};
    return b;
  }

  build() {
    return new ConditionalFormatRuleMock(this._realm, this._spec);
  }

  toString() {
    return 'ConditionalFormatRuleBuilder';
  }
}

/** FilterCriteria (immutable, recorded only). */
class FilterCriteriaMock {
  constructor(realm, spec) {
    defineHidden(this, {_realm: realm, _spec: {...spec}});
  }

  getHiddenValues() {
    return this._realm.toContext([...(this._spec.hiddenValues || [])]);
  }

  getVisibleValues() {
    return this._realm.toContext([...(this._spec.visibleValues || [])]);
  }

  getCriteriaType() {
    return this._spec.criteria || null;
  }

  getCriteriaValues() {
    return this._realm.toContext([...(this._spec.values || [])]);
  }

  copy() {
    const b = new FilterCriteriaBuilderMock(this._realm);
    b._spec = {...this._spec};
    return b;
  }

  describe() {
    return {...this._spec};
  }

  toString() {
    return 'FilterCriteria';
  }
}

/** FilterCriteriaBuilder (SpreadsheetApp.newFilterCriteria()). */
class FilterCriteriaBuilderMock {
  constructor(realm) {
    defineHidden(this, {_realm: realm, _spec: {}});
  }

  setHiddenValues(values) {
    this._spec.hiddenValues = this._realm.toHost(values).slice();
    return this;
  }

  setVisibleValues(values) {
    this._spec.visibleValues = this._realm.toHost(values).slice();
    return this;
  }

  whenCellEmpty() {
    Object.assign(this._spec, {criteria: 'CELL_EMPTY', values: []});
    return this;
  }

  whenCellNotEmpty() {
    Object.assign(this._spec, {criteria: 'CELL_NOT_EMPTY', values: []});
    return this;
  }

  whenTextEqualTo(t) {
    Object.assign(this._spec, {criteria: 'TEXT_EQUAL_TO', values: [t]});
    return this;
  }

  whenTextContains(t) {
    Object.assign(this._spec, {criteria: 'TEXT_CONTAINS', values: [t]});
    return this;
  }

  whenFormulaSatisfied(f) {
    Object.assign(this._spec, {criteria: 'CUSTOM_FORMULA', values: [f]});
    return this;
  }

  build() {
    return new FilterCriteriaMock(this._realm, this._spec);
  }

  toString() {
    return 'FilterCriteriaBuilder';
  }
}

// ---------------------------------------------------------------------------
// Protection and named ranges
// ---------------------------------------------------------------------------

/** Protection (recorded only; never enforced). */
class ProtectionMock {
  /**
   * @param {!Object} sheet SheetMock owning the protection.
   * @param {string} type 'RANGE' | 'SHEET'
   * @param {?Object} range RangeMock for RANGE protections.
   */
  constructor(sheet, type, range) {
    defineHidden(this, {
      _sheet: sheet, _type: type, _range: range, _description: '', _warningOnly: false,
      _editors: [], _domainEdit: true, _unprotected: [],
    });
  }

  getProtectionType() {
    return this._type;
  }

  getRange() {
    return this._range || this._sheet.getDataRange();
  }

  setDescription(d) {
    this._description = String(d);
    return this;
  }

  getDescription() {
    return this._description;
  }

  setWarningOnly(b) {
    this._warningOnly = Boolean(b);
    return this;
  }

  isWarningOnly() {
    return this._warningOnly;
  }

  addEditor(e) {
    this._editors.push(String(e && e.getEmail ? e.getEmail() : e));
    return this;
  }

  addEditors(list) {
    for (const e of list) this.addEditor(e);
    return this;
  }

  removeEditor(e) {
    const email = String(e && e.getEmail ? e.getEmail() : e);
    this._editors = this._editors.filter((x) => x !== email);
    return this;
  }

  removeEditors(list) {
    for (const e of list) this.removeEditor(e);
    return this;
  }

  getEditors() {
    return this._sheet._env.realm.toContext([]);
  }

  canEdit() {
    return true;
  }

  setDomainEdit(b) {
    this._domainEdit = Boolean(b);
    return this;
  }

  canDomainEdit() {
    return this._domainEdit;
  }

  setUnprotectedRanges(ranges) {
    this._unprotected = [...ranges];
    return this;
  }

  getUnprotectedRanges() {
    return this._sheet._env.realm.toContext([...this._unprotected]);
  }

  remove() {
    this._sheet._protections = this._sheet._protections.filter((p) => p !== this);
  }

  describe() {
    return {
      type: this._type,
      range: this._range ? this._range.getA1Notation() : null,
      description: this._description,
      warningOnly: this._warningOnly,
    };
  }

  toString() {
    return 'Protection';
  }
}

/** NamedRange. */
class NamedRangeMock {
  /**
   * @param {!Object} spreadsheet SpreadsheetMock.
   * @param {string} name
   * @param {!Object} range RangeMock.
   */
  constructor(spreadsheet, name, range) {
    defineHidden(this, {_ss: spreadsheet, _name: name, _range: range});
  }

  getName() {
    return this._name;
  }

  setName(name) {
    this._ss._namedRanges.delete(this._name);
    this._name = String(name);
    this._ss._namedRanges.set(this._name, this);
    return this;
  }

  getRange() {
    return this._range;
  }

  setRange(range) {
    this._range = range;
    return this;
  }

  remove() {
    this._ss._namedRanges.delete(this._name);
  }

  toString() {
    return 'NamedRange';
  }
}

module.exports = {
  ENUMS,
  enumOf,
  normalizeColor,
  DataValidationMock,
  DataValidationBuilderMock,
  TextStyleMock,
  TextStyleBuilderMock,
  RichTextValueMock,
  RichTextValueBuilderMock,
  ConditionalFormatRuleMock,
  ConditionalFormatRuleBuilderMock,
  FilterCriteriaMock,
  FilterCriteriaBuilderMock,
  ProtectionMock,
  NamedRangeMock,
};
