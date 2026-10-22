/**
 * @fileoverview Helpers for moving data across the node:vm realm boundary.
 *
 * Code under test runs inside a `node:vm` context, which has its own set of
 * JavaScript intrinsics (`Array`, `Object`, `Date`, `Error`, ...). Objects that
 * are created by the (host-side) mocks would otherwise fail `instanceof` checks
 * inside the Apps Script code (`value instanceof Date` is a very common Apps
 * Script idiom for spreadsheet values), and objects coming *out* of the context
 * fail `assert.deepStrictEqual` against host literals because their prototypes
 * differ. The helpers in this file convert plain data (arrays, plain objects,
 * Dates) in both directions while leaving everything else (mock class
 * instances, functions, primitives) untouched.
 */
'use strict';

const vm = require('node:vm');

/**
 * Returns true when `value` is a Date from any realm.
 * @param {*} value
 * @return {boolean}
 */
function isDate(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

/**
 * Returns true when `value` is an Error-like object from any realm.
 * @param {*} value
 * @return {boolean}
 */
function isError(value) {
  return Object.prototype.toString.call(value) === '[object Error]';
}

/**
 * Creates realm helpers bound to a vm context.
 * @param {!Object} context A contextified object (from vm.createContext).
 * @return {!Realm}
 */
function createRealm(context) {
  /** @type {{Object: Function, Array: Function, Date: Function, Error: Function, TypeError: Function}} */
  const intrinsics = vm.runInContext(
    '({Object: Object, Array: Array, Date: Date, Error: Error, TypeError: TypeError})',
    context);
  const ctxObjectProto = intrinsics.Object.prototype;
  const ctxArrayProto = intrinsics.Array.prototype;

  /**
   * Deep-converts host plain data into context-realm plain data.
   * Context-realm objects and non-plain objects are returned as-is.
   * @param {*} value
   * @param {!Map=} seen Cycle tracking.
   * @return {*}
   */
  function toContext(value, seen) {
    if (value === null || typeof value !== 'object') return value;
    if (!(seen instanceof Map)) seen = new Map();  // Tolerate use as an Array#map callback.
    if (seen.has(value)) return seen.get(value);
    const proto = Object.getPrototypeOf(value);
    if (isDate(value)) {
      return proto === intrinsics.Date.prototype ? value : new intrinsics.Date(value.getTime());
    }
    if (Array.isArray(value)) {
      if (proto === ctxArrayProto) return value;
      const out = new intrinsics.Array();
      seen.set(value, out);
      for (let i = 0; i < value.length; i++) out[i] = toContext(value[i], seen);
      return out;
    }
    if (proto === Object.prototype || proto === null) {
      const out = new intrinsics.Object();
      seen.set(value, out);
      for (const key of Object.keys(value)) out[key] = toContext(value[key], seen);
      return out;
    }
    return value;
  }

  /**
   * Deep-converts context-realm plain data into host-realm plain data.
   * Host-realm objects and non-plain objects (mocks, class instances) are
   * returned as-is.
   * @param {*} value
   * @param {!Map=} seen Cycle tracking.
   * @return {*}
   */
  function toHost(value, seen) {
    if (value === null || typeof value !== 'object') return value;
    if (!(seen instanceof Map)) seen = new Map();  // Tolerate use as an Array#map callback.
    if (seen.has(value)) return seen.get(value);
    const proto = Object.getPrototypeOf(value);
    if (isDate(value)) {
      return proto === Date.prototype ? value : new Date(value.getTime());
    }
    if (Array.isArray(value)) {
      const out = [];
      seen.set(value, out);
      for (let i = 0; i < value.length; i++) out[i] = toHost(value[i], seen);
      return out;
    }
    if (proto === ctxObjectProto || proto === Object.prototype || proto === null) {
      const out = {};
      seen.set(value, out);
      for (const key of Object.keys(value)) out[key] = toHost(value[key], seen);
      return out;
    }
    return value;
  }

  /**
   * Creates an Error in the context realm, named like Apps Script service
   * exceptions (`Exception: <message>` when stringified).
   * @param {string} message
   * @param {string=} name Defaults to 'Exception'.
   * @param {!Object=} extra Extra own properties to attach (e.g. `details`).
   * @return {!Error}
   */
  function error(message, name = 'Exception', extra = undefined) {
    const err = new intrinsics.Error(message);
    Object.defineProperty(err, 'name', {value: name, writable: true, configurable: true});
    // Re-derive the stack header so it reflects the custom name.
    if (typeof err.stack === 'string') {
      err.stack = err.stack.replace(/^Error/, name);
    }
    if (extra) Object.assign(err, toContext(extra));
    return err;
  }

  /**
   * Creates a context-realm Date.
   * @param {number|!Date} time
   * @return {!Date}
   */
  function date(time) {
    return new intrinsics.Date(isDate(time) ? time.getTime() : time);
  }

  return {
    context,
    intrinsics,
    toContext,
    toHost,
    error,
    date,
    isDate,
    isError,
    /** @return {!Array} A new, empty context-realm array. */
    array() {
      return new intrinsics.Array();
    },
    /** @return {!Object} A new, empty context-realm object. */
    object() {
      return new intrinsics.Object();
    },
  };
}

/**
 * @typedef {{
 *   context: !Object,
 *   intrinsics: !Object,
 *   toContext: function(*): *,
 *   toHost: function(*): *,
 *   error: function(string, string=, !Object=): !Error,
 *   date: function((number|!Date)): !Date,
 *   isDate: function(*): boolean,
 *   isError: function(*): boolean,
 *   array: function(): !Array,
 *   object: function(): !Object,
 * }}
 */
let Realm;  // eslint-disable-line no-unused-vars

module.exports = {createRealm, isDate, isError};
