/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Generic helpers: value conversion, list and table text
 * formats, hashing and GTM parameter manipulation.
 */

/**
 * An error whose message is meant to be shown to the user as-is.
 */
class UserError extends Error {
  /** @param {string} message */
  constructor(message) {
    super(message);
    this.name = 'UserError';
  }
}

/**
 * Returns true if a cell value is empty.
 * @param {*} value
 * @return {boolean}
 */
function isBlank(value) {
  return value === '' || value === null || value === undefined ||
      (typeof value === 'string' && value.trim() === '');
}

/**
 * Converts a cell value to a boolean.
 * @param {*} value
 * @return {boolean}
 */
function toBool(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  const text = String(value === null || value === undefined ? '' : value)
      .trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === 'y' || text === '1';
}

/**
 * Pads a number with a leading zero.
 * @param {number} n
 * @return {string}
 */
function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

/**
 * Converts a cell value to a trimmed string with normalised line breaks.
 * Trailing white space is removed from every line.
 * @param {*} value
 * @return {string}
 */
function toText(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    const date = `${value.getFullYear()}-${pad2_(value.getMonth() + 1)}-` +
        `${pad2_(value.getDate())}`;
    if (value.getHours() === 0 && value.getMinutes() === 0) return date;
    return `${date} ${pad2_(value.getHours())}:${pad2_(value.getMinutes())}`;
  }
  return String(value)
      .replace(/\r\n?/g, '\n')
      .split('\n')
      .map(line => line.replace(/\s+$/, ''))
      .join('\n')
      .trim();
}

/**
 * Parses a list typed in a cell. Items are separated by new lines or, if the
 * text has no line breaks, by commas.
 * @param {*} value
 * @return {!Array<string>}
 */
function parseList(value) {
  const text = toText(value);
  if (!text) return [];
  const parts = text.indexOf('\n') !== -1 ? text.split('\n') : text.split(',');
  return parts.map(part => part.trim()).filter(part => part !== '');
}

/**
 * Formats a list for a cell. Items are separated by commas unless an item
 * contains a comma, in which case every item is written on its own line.
 * @param {!Array<string>} items
 * @return {string}
 */
function formatList(items) {
  const clean = (items || []).map(item => toText(item).replace(/\n/g, ' '))
      .filter(item => item !== '');
  if (clean.some(item => item.indexOf(',') !== -1)) return clean.join('\n');
  return clean.join(', ');
}

/**
 * Parses a two-column table typed in a cell. Every line is one row in the
 * "name = value" format.
 * @param {*} value
 * @return {!Array<{name: string, value: string}>}
 */
function parseTable(value) {
  const text = toText(value);
  if (!text) return [];
  return text.split('\n')
      .map(line => line.trim())
      .filter(line => line !== '')
      .map(line => {
        let index = line.indexOf(' = ');
        let separatorLength = 3;
        if (index === -1) {
          index = line.indexOf('=');
          separatorLength = 1;
        }
        if (index === -1) return {name: line, value: ''};
        return {
          name: line.slice(0, index).trim(),
          value: line.slice(index + separatorLength).trim()
        };
      });
}

/**
 * Formats a two-column table for a cell ("name = value" on every line).
 * @param {!Array<{name: string, value: string}>} rows
 * @return {string}
 */
function formatTable(rows) {
  return (rows || []).map(row => {
    const name = toText(row.name).replace(/\n/g, ' ');
    const value = toText(row.value).replace(/\n/g, ' ');
    return `${name} = ${value}`;
  }).join('\n');
}

/**
 * Returns a short, stable hash of a string (FNV-1a, 32 bit).
 * @param {string} text
 * @return {string}
 */
function hashString(text) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return ('0000000' + hash.toString(16)).slice(-8);
}

/**
 * Deep copies a JSON-compatible object.
 * @param {T} value
 * @return {T}
 * @template T
 */
function deepCopy(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

/**
 * Splits an array into chunks.
 * @param {!Array<T>} items
 * @param {number} size
 * @return {!Array<!Array<T>>}
 * @template T
 */
function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

/**
 * Shortens a text for display.
 * @param {string} text
 * @param {number} maxLength
 * @return {string}
 */
function truncateText(text, maxLength) {
  const value = String(text);
  return value.length > maxLength ?
      value.slice(0, Math.max(0, maxLength - 1)) + '…' : value;
}

/**
 * Returns a value that fits in a cell: long texts are truncated.
 * @param {*} value
 * @return {*}
 */
function cellSafeValue(value) {
  if (typeof value === 'string' && value.length > LIMITS.maxCellLength) {
    return value.slice(0, LIMITS.maxCellLength) + TRUNCATION_SUFFIX;
  }
  return value;
}

/**
 * Compares two strings case-insensitively, for sorting.
 * @param {string} a
 * @param {string} b
 * @return {number}
 */
function compareText(a, b) {
  return String(a).toLowerCase().localeCompare(String(b).toLowerCase());
}

// ---------------------------------------------------------------------------
// GTM parameter helpers
// ---------------------------------------------------------------------------

/**
 * Finds a parameter by key.
 * @param {?Array<!Object>|undefined} parameters
 * @param {string} key
 * @return {!Object|undefined}
 */
function findParam(parameters, key) {
  return (parameters || []).find(param => param.key === key);
}

/**
 * Returns the value of a scalar parameter as a string ('' if missing).
 * @param {!Object} entity A GTM entity with a parameter array.
 * @param {string} key
 * @return {string}
 */
function getParamValue(entity, key) {
  const param = findParam(entity.parameter, key);
  return param && param.value !== undefined && param.value !== null ?
      String(param.value) : '';
}

/**
 * Adds a parameter or replaces the existing parameter with the same key.
 * @param {!Object} entity
 * @param {!Object} param
 */
function putParam(entity, param) {
  entity.parameter = entity.parameter || [];
  const index = entity.parameter.findIndex(p => p.key === param.key);
  if (index === -1) {
    entity.parameter.push(param);
  } else {
    entity.parameter[index] = param;
  }
}

/**
 * Removes a parameter.
 * @param {!Object} entity
 * @param {string} key
 */
function removeParam(entity, key) {
  if (entity.parameter) {
    entity.parameter = entity.parameter.filter(param => param.key !== key);
  }
}

/**
 * Sets a scalar parameter. An empty value removes the parameter unless
 * options.keepEmpty is set. The type of an existing parameter is kept when it
 * is compatible with the value.
 * @param {!Object} entity
 * @param {string} key
 * @param {*} value
 * @param {{type: (string|undefined), keepEmpty: (boolean|undefined)}=} options
 */
function setParamValue(entity, key, value, options) {
  const opts = options || {};
  const text = toText(value);
  if (text === '' && !opts.keepEmpty) {
    removeParam(entity, key);
    return;
  }
  const existing = findParam(entity.parameter, key);
  let type = opts.type || (existing && existing.type) || 'template';
  if (type === 'boolean' && !/^(true|false)$/i.test(text)) type = 'template';
  if (type === 'integer' && !/^-?\d+$/.test(text)) type = 'template';
  if (type === 'list' || type === 'map') type = 'template';
  putParam(entity, {
    type: type,
    key: key,
    value: type === 'boolean' ? text.toLowerCase() : text
  });
}

/**
 * Returns the value of a boolean parameter.
 * @param {!Object} entity
 * @param {string} key
 * @param {boolean=} defaultValue Value used if the parameter is missing.
 * @return {boolean}
 */
function getBoolParam(entity, key, defaultValue) {
  const param = findParam(entity.parameter, key);
  if (!param || param.value === undefined) return Boolean(defaultValue);
  return String(param.value).toLowerCase() === 'true';
}

/**
 * Sets a boolean parameter.
 * @param {!Object} entity
 * @param {string} key
 * @param {*} value
 */
function setBoolParam(entity, key, value) {
  putParam(entity, {
    type: 'boolean',
    key: key,
    value: toBool(value) ? 'true' : 'false'
  });
}

/**
 * Returns the values of a list parameter whose items are scalars.
 * @param {!Object} entity
 * @param {string} key
 * @return {!Array<string>}
 */
function getListParamValues(entity, key) {
  const param = findParam(entity.parameter, key);
  if (!param || !param.list) return [];
  return param.list.map(item =>
    item.value !== undefined && item.value !== null ? String(item.value) : '');
}

/**
 * Sets a list parameter with scalar items. An empty list removes the
 * parameter.
 * @param {!Object} entity
 * @param {string} key
 * @param {!Array<string>} values
 */
function setListParamValues(entity, key, values) {
  if (!values.length) {
    removeParam(entity, key);
    return;
  }
  putParam(entity, {
    type: 'list',
    key: key,
    list: values.map(value => ({type: 'template', value: String(value)}))
  });
}

/**
 * Returns a field value from a map parameter item.
 * @param {!Object} mapItem A parameter of type map.
 * @param {string} field
 * @return {string}
 */
function getMapField(mapItem, field) {
  const entry = (mapItem.map || []).find(e => e.key === field);
  return entry && entry.value !== undefined && entry.value !== null ?
      String(entry.value) : '';
}

/**
 * Returns the rows of a table parameter (a list of maps).
 * @param {!Object} entity
 * @param {string} key
 * @param {!Array<string>} fields Map keys used for the name and the value
 *     (the second field is optional).
 * @return {!Array<{name: string, value: string, item: !Object}>}
 */
function getTableRows(entity, key, fields) {
  const param = findParam(entity.parameter, key);
  if (!param || !param.list) return [];
  return param.list.map(item => ({
    name: getMapField(item, fields[0]),
    value: fields[1] ? getMapField(item, fields[1]) : '',
    item: item
  }));
}

/**
 * Builds a map item of a table parameter. Existing map entries are kept so
 * that additional fields of a row are not lost.
 * @param {!Array<string>} fields
 * @param {{name: string, value: string}} row
 * @param {!Object=} existingItem
 * @return {!Object}
 */
function buildTableItem(fields, row, existingItem) {
  const item = existingItem ? deepCopy(existingItem) :
      {type: 'map', map: []};
  item.type = 'map';
  item.map = item.map || [];
  const values = [row.name, row.value];
  fields.forEach((field, index) => {
    if (!field) return;
    const entry = item.map.find(e => e.key === field);
    if (entry) {
      entry.value = String(values[index]);
      if (entry.type !== 'boolean' && entry.type !== 'integer') {
        entry.type = 'template';
      }
    } else {
      item.map.push({type: 'template', key: field,
                     value: String(values[index])});
    }
  });
  return item;
}

/**
 * Replaces the rows of a table parameter. Existing rows with the same name
 * are reused so that their additional fields are kept. An empty list of rows
 * removes the parameter.
 * @param {!Object} entity
 * @param {string} key
 * @param {!Array<string>} fields
 * @param {!Array<{name: string, value: string}>} rows
 */
function setTableRows(entity, key, fields, rows) {
  if (!rows.length) {
    removeParam(entity, key);
    return;
  }
  const available = getTableRows(entity, key, fields);
  const list = rows.map(row => {
    const index = available.findIndex(existing => existing.name === row.name);
    let existingItem;
    if (index !== -1) {
      existingItem = available[index].item;
      available.splice(index, 1);
    }
    return buildTableItem(fields, row, existingItem);
  });
  putParam(entity, {type: 'list', key: key, list: list});
}

/**
 * Converts a parameter to a plain JavaScript value: scalars become strings,
 * lists become arrays and maps become objects.
 * @param {!Object} param
 * @return {*}
 */
function paramToSimple(param) {
  if (param.type === 'list') return (param.list || []).map(paramToSimple);
  if (param.type === 'map') {
    const object = {};
    (param.map || []).forEach(entry => {
      object[entry.key] = paramToSimple(entry);
    });
    return object;
  }
  return param.value !== undefined && param.value !== null ?
      String(param.value) : '';
}

/**
 * Converts a parameter to the text shown in a cell. Lists and maps are shown
 * as compact JSON.
 * @param {!Object} param
 * @return {string}
 */
function paramToText(param) {
  if (param.type === 'list' || param.type === 'map') {
    return JSON.stringify(paramToSimple(param));
  }
  return paramToSimple(param);
}

/**
 * Converts a plain JavaScript value back to a parameter. Types are taken from
 * the original parameter when possible.
 * @param {string|undefined} key
 * @param {*} simple
 * @param {!Object=} original The original parameter at the same position.
 * @return {!Object}
 */
function simpleToParam(key, simple, original) {
  let param;
  if (Array.isArray(simple)) {
    const originalList = original && original.type === 'list' ?
        original.list || [] : [];
    param = {
      type: 'list',
      list: simple.map((item, index) => simpleToParam(
          undefined, item, originalList[index] || originalList[0]))
    };
  } else if (simple !== null && typeof simple === 'object') {
    const originalMap = original && original.type === 'map' ?
        original.map || [] : [];
    param = {
      type: 'map',
      map: Object.keys(simple).map(mapKey => simpleToParam(
          mapKey, simple[mapKey], originalMap.find(e => e.key === mapKey)))
    };
  } else {
    const value = simple === null || simple === undefined ? '' :
        String(simple);
    const scalarTypes = ['template', 'integer', 'boolean', 'triggerReference',
                         'tagReference'];
    let type = original && scalarTypes.indexOf(original.type) !== -1 ?
        original.type : (typeof simple === 'boolean' ? 'boolean' : 'template');
    if (type === 'boolean' && !/^(true|false)$/i.test(value)) type = 'template';
    if (type === 'integer' && !/^-?\d+$/.test(value)) type = 'template';
    param = {type: type, value: type === 'boolean' ? value.toLowerCase() :
        value};
    if (original && original.isWeakReference !== undefined) {
      param.isWeakReference = original.isWeakReference;
    }
  }
  if (key !== undefined) param.key = key;
  return param;
}

/**
 * Converts the text of a cell to a parameter. JSON is expected for lists and
 * maps.
 * @param {string} key
 * @param {*} value
 * @param {!Object=} original The current parameter, if any.
 * @return {!Object}
 */
function textToParam(key, value, original) {
  const text = toText(value);
  const isStructured = original &&
      (original.type === 'list' || original.type === 'map');
  if (isStructured || /^[\[{]/.test(text)) {
    let parsed;
    try {
      parsed = JSON.parse(text || (original && original.type === 'map' ?
          '{}' : '[]'));
    } catch (e) {
      throw new UserError(`The value of "${key}" must be valid JSON ` +
          `(for example ["a","b"] or [{"name":"x","value":"y"}]).`);
    }
    return simpleToParam(key, parsed, original);
  }
  return simpleToParam(key, text, original);
}

/**
 * Returns a copy of an entity without the fields that identify it, so that
 * it can be used to create a new entity.
 * @param {!Object} entity
 * @param {string} idField
 * @return {!Object}
 */
function stripEntityIdentity(entity, idField) {
  const copy = deepCopy(entity);
  ['path', 'accountId', 'containerId', 'workspaceId', 'fingerprint',
   'tagManagerUrl', idField].forEach(field => {
    delete copy[field];
  });
  return copy;
}
