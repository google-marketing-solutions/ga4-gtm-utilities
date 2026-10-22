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
 * @fileoverview Builders for the columns used by the feature sheets.
 *
 * A column definition has:
 * - key: unique key within the sheet.
 * - header, group, note, width: layout.
 * - kind: 'text' (default), 'multiline', 'checkbox', 'list' (comma or new
 *   line separated values) or 'table' ("name = value" per line).
 * - dropdown: an array of values, or the name of a dynamic list ('Triggers',
 *   'Folders', 'Variables', 'Tags').
 * - readOnly, createOnly, required: editing rules.
 * - appliesTo(entity, data, ctx): whether the column applies to an entity.
 * - get(entity, data, ctx): the cell value.
 * - set(draft, value, data, ctx): applies a cell value to an entity.
 */

/**
 * Returns an appliesTo function for a list of entity types.
 * @param {!Array<string>|undefined} types
 * @return {function(!Object): boolean|undefined}
 */
function appliesToTypes_(types) {
  if (!types) return undefined;
  return entity => types.indexOf(entity.type) !== -1;
}

/**
 * Builds a column with common options.
 * @param {!Object} base
 * @param {!Object=} options Overrides (header, group, note, width, types,
 *     required, readOnly, createOnly, dropdown).
 * @return {!Object}
 */
function buildColumn_(base, options) {
  const opts = options || {};
  const col = Object.assign({}, base);
  ['key', 'header', 'group', 'note', 'width', 'required', 'readOnly',
   'createOnly', 'dropdown', 'strictDropdown', 'kind'].forEach(field => {
    if (opts[field] !== undefined) col[field] = opts[field];
  });
  if (opts.types) col.appliesTo = appliesToTypes_(opts.types);
  if (opts.appliesTo) col.appliesTo = opts.appliesTo;
  return col;
}

/** Column builders. */
const COLUMNS = {
  /**
   * Entity name.
   * @param {string} label E.g. "Tag".
   * @param {!Object=} options
   * @return {!Object}
   */
  name(label, options) {
    return buildColumn_({
      key: 'name',
      header: `${label} Name`,
      width: 280,
      required: true,
      note: 'Name of the item in GTM (must be unique).',
      get: entity => entity.name || '',
      set: (draft, value) => {
        const name = toText(value);
        if (!name) throw new UserError('The name can\'t be empty.');
        draft.name = name;
      }
    }, options);
  },

  /**
   * Entity type, chosen from a list of readable names.
   * @param {!Object<string, string>} names Map of type code to name.
   * @param {string} label E.g. "Tag".
   * @param {!Array<string>} allowedTypes Types that can be created here.
   * @param {!Object=} options
   * @return {!Object}
   */
  type(names, label, allowedTypes, options) {
    return buildColumn_({
      key: 'type',
      header: `${label} Type`,
      width: 220,
      createOnly: true,
      required: true,
      dropdown: allowedTypes.map(type => names[type] || type),
      note: `Type of the ${label.toLowerCase()}. Choose one of the values ` +
          'in the dropdown.',
      get: (entity, data) => getTypeLabel(names, entity.type, data),
      set: (draft, value, data) => {
        const type = resolveTypeLabel(names, value, data);
        if (!type) throw new UserError('The type is required.');
        if (allowedTypes.length && allowedTypes.indexOf(type) === -1) {
          throw new UserError(`"${toText(value)}" is not a type that can be ` +
              `created in this sheet. Use one of: ` +
              allowedTypes.map(t => names[t] || t).join(', ') + '.');
        }
        draft.type = type;
      }
    }, options);
  },

  /**
   * Read-only entity type.
   * @param {!Object<string, string>} names
   * @param {string} label
   * @param {!Object=} options
   * @return {!Object}
   */
  typeReadOnly(names, label, options) {
    return buildColumn_({
      key: 'type',
      header: `${label} Type`,
      width: 220,
      readOnly: true,
      get: (entity, data) => getTypeLabel(names, entity.type, data)
    }, options);
  },

  /**
   * Folder of the entity. Unknown folder names are created automatically.
   * @param {!Object=} options
   * @return {!Object}
   */
  folder(options) {
    return buildColumn_({
      key: 'folder',
      header: 'Folder',
      width: 160,
      dropdown: 'Folders',
      note: 'Folder name. A folder that does not exist yet is created ' +
          'automatically. Leave empty for no folder.',
      get: (entity, data) => data.getFolderName(entity.parentFolderId),
      set: (draft, value, data) => {
        const id = data.getFolderIdForName(value);
        if (id) {
          draft.parentFolderId = id;
        } else {
          delete draft.parentFolderId;
        }
      }
    }, options);
  },

  /**
   * Notes of the entity.
   * @param {!Object=} options
   * @return {!Object}
   */
  notes(options) {
    return buildColumn_({
      key: 'notes',
      header: 'Notes',
      width: 220,
      kind: 'multiline',
      get: entity => entity.notes || '',
      set: (draft, value) => {
        const notes = toText(value);
        if (notes) {
          draft.notes = notes;
        } else {
          delete draft.notes;
        }
      }
    }, options);
  },

  /**
   * Paused state of a tag or client.
   * @param {!Object=} options
   * @return {!Object}
   */
  paused(options) {
    return buildColumn_({
      key: 'paused',
      header: 'Paused',
      width: 70,
      kind: 'checkbox',
      note: 'Tick to pause.',
      get: entity => Boolean(entity.paused),
      set: (draft, value) => {
        draft.paused = toBool(value);
      }
    }, options);
  },

  /**
   * Firing or blocking triggers of a tag.
   * @param {string} field 'firingTriggerId' or 'blockingTriggerId'.
   * @param {!Object=} options
   * @return {!Object}
   */
  triggers(field, options) {
    const isFiring = field === 'firingTriggerId';
    return buildColumn_({
      key: field,
      header: isFiring ? 'Firing Triggers' : 'Exception (Blocking) Triggers',
      width: 240,
      kind: 'list',
      dropdown: 'Triggers',
      note: 'Trigger names, separated by commas (or one per line if a name ' +
          'contains a comma). Built-in triggers: All Pages, ' +
          'Initialization - All Pages, Consent Initialization - All Pages.',
      get: (entity, data) => data.triggerIdsToText(entity[field]),
      set: (draft, value, data) => {
        const ids = data.triggerTextToIds(value);
        if (ids.length) {
          draft[field] = ids;
        } else {
          delete draft[field];
        }
      }
    }, options);
  },

  /**
   * A scalar parameter.
   * @param {string} paramKey
   * @param {string} header
   * @param {!Object=} options paramType: type for new parameters
   *     (default template).
   * @return {!Object}
   */
  param(paramKey, header, options) {
    const opts = options || {};
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 180,
      get: entity => getParamValue(entity, paramKey),
      set: (draft, value) => {
        setParamValue(draft, paramKey, value, {type: opts.paramType});
      }
    }, options);
  },

  /**
   * A parameter whose values have readable labels (shown as a dropdown).
   * @param {string} paramKey
   * @param {string} header
   * @param {!Object<string, string>} labels Map of API value to label.
   * @param {!Object=} options defaultValue: value used if the parameter is
   *     missing.
   * @return {!Object}
   */
  selectParam(paramKey, header, labels, options) {
    const opts = options || {};
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 170,
      dropdown: Object.keys(labels).map(value => labels[value]),
      get: entity => {
        const value = getParamValue(entity, paramKey) ||
            (opts.defaultValue || '');
        return labels[value] || value;
      },
      set: (draft, value) => {
        const text = toText(value);
        const lower = text.toLowerCase();
        const apiValue = Object.keys(labels).find(key =>
          labels[key].toLowerCase() === lower || key.toLowerCase() === lower);
        setParamValue(draft, paramKey, apiValue !== undefined ? apiValue :
            text, {type: opts.paramType});
      }
    }, options);
  },

  /**
   * A boolean parameter, shown as a checkbox.
   * @param {string} paramKey
   * @param {string} header
   * @param {!Object=} options defaultValue: value used if the parameter is
   *     missing.
   * @return {!Object}
   */
  boolParam(paramKey, header, options) {
    const opts = options || {};
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 110,
      kind: 'checkbox',
      get: entity => getBoolParam(entity, paramKey, opts.defaultValue),
      set: (draft, value) => {
        setBoolParam(draft, paramKey, value);
      }
    }, options);
  },

  /**
   * A list parameter with scalar values.
   * @param {string} paramKey
   * @param {string} header
   * @param {!Object=} options
   * @return {!Object}
   */
  listParam(paramKey, header, options) {
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 220,
      kind: 'list',
      get: entity => formatList(getListParamValues(entity, paramKey)),
      set: (draft, value) => {
        setListParamValues(draft, paramKey, parseList(value));
      }
    }, options);
  },

  /**
   * A table parameter with a name and a value per row ("name = value" per
   * line in the cell).
   * @param {string} paramKey
   * @param {string} header
   * @param {!Array<string>} fields Map keys of the name and the value.
   * @param {!Object=} options
   * @return {!Object}
   */
  tableParam(paramKey, header, fields, options) {
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 280,
      kind: 'table',
      note: 'One row per line, in the format: name = value',
      get: entity => formatTable(getTableRows(entity, paramKey, fields)),
      set: (draft, value) => {
        const rows = parseTable(value);
        rows.forEach(row => {
          if (!row.name) {
            throw new UserError(`${header}: every line needs a name ` +
                '("name = value").');
          }
        });
        setTableRows(draft, paramKey, fields, rows);
      }
    }, options);
  },

  /**
   * A table parameter with a single field per row, shown as a list.
   * @param {string} paramKey
   * @param {string} header
   * @param {string} field Map key of the value.
   * @param {!Object=} options
   * @return {!Object}
   */
  singleFieldTableParam(paramKey, header, field, options) {
    return buildColumn_({
      key: 'p:' + paramKey,
      header: header,
      width: 220,
      kind: 'list',
      get: entity => formatList(
          getTableRows(entity, paramKey, [field]).map(row => row.name)),
      set: (draft, value) => {
        setTableRows(draft, paramKey, [field],
            parseList(value).map(name => ({name: name, value: ''})));
      }
    }, options);
  }
};
