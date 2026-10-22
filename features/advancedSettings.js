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
 * @fileoverview The "Advanced Settings" sheet: one row per parameter of any
 * tag, variable, client or transformation. This sheet makes every setting
 * editable, including the settings of non-Google tags and custom templates.
 */

/**
 * Returns the definition of the Advanced Settings sheet.
 * @return {!Object}
 */
function advancedSettingsSheetDef_() {
  const kinds = ['tag', 'variable', 'client', 'transformation'];
  return {
    key: 'advancedSettings',
    sheetName: 'Advanced Settings',
    title: 'ADVANCED SETTINGS (ALL PARAMETERS)',
    tabColor: TAB_COLORS.general,
    mode: 'item',
    kinds: kinds,
    contexts: ['web', 'server'],
    itemLabel: 'setting',
    itemLabelPlural: 'settings',
    idHeader: 'Item ID',
    nameColumn: 'itemName',
    columns: [
      {key: 'itemType', header: 'Item Type', group: 'ITEM', width: 100,
       createOnly: true,
       dropdown: ['Tag', 'Variable', 'Client', 'Transformation'],
       note: 'Tag, Variable, Client or Transformation.'},
      {key: 'itemName', header: 'Item Name', group: 'ITEM', width: 260,
       createOnly: true,
       note: 'Name of the tag, variable, client or transformation. For new ' +
           'rows, the name is used to find the item.'},
      {key: 'itemKind', header: 'Item Type Detail', group: 'ITEM',
       width: 200, readOnly: true},
      {key: 'param', header: 'Parameter Key', group: 'PARAMETER', width: 220,
       note: 'Key of the parameter as used by the Tag Manager API (e.g. ' +
           'html, eventName, pixelId).'},
      {key: 'value', header: 'Value', group: 'PARAMETER', width: 360,
       kind: 'multiline',
       note: 'Value of the parameter. Lists and tables are shown as JSON, ' +
           'e.g. ["a","b"] or [{"name":"x","value":"y"}]; keep the JSON ' +
           'format when editing them.'},
      {key: 'valueType', header: 'Value Type', group: 'PARAMETER', width: 90,
       readOnly: true,
       note: 'template (text), boolean, integer, list or map.'}
    ],
    isEligible: entity => (entity.parameter || []).length > 0,
    explode: (entity, kind, data) => explodeParamItems_(entity, kind, {
      itemType: ITEM_TYPE_LABELS[kind],
      itemName: entity.name,
      itemKind: getTypeLabel(getEntityTypeNames_(kind), entity.type, data)
    }),
    findEntity: (row, data, ctx, op) => findItemEntity_(row, data, kinds,
        () => true, op),
    updateItem: (draft, row) => {
      checkItemNameUnchanged_(draft, row);
      return updateParamItem_(draft, row);
    },
    createItem: (draft, row, data, ctx, kind) =>
      createParamItem_(draft, row, kind),
    deleteItem: (draft, row) => deleteParamItem_(draft, row)
  };
}

/**
 * Returns the type names of an entity kind.
 * @param {string} kind
 * @return {!Object<string, string>}
 */
function getEntityTypeNames_(kind) {
  return {
    tag: TAG_TYPE_NAMES,
    variable: VARIABLE_TYPE_NAMES,
    client: CLIENT_TYPE_NAMES,
    transformation: TRANSFORMATION_TYPE_NAMES,
    gtagConfig: GTAG_CONFIG_TYPE_NAMES
  }[kind] || {};
}

/**
 * Builds one row per parameter of an entity.
 * @param {!Object} entity
 * @param {string} kind
 * @param {!Object} base Values of the item columns.
 * @return {!Array<!Object>}
 */
function explodeParamItems_(entity, kind, base) {
  return (entity.parameter || []).map(param => {
    const fullText = paramToText(param);
    const text = cellSafeValue(fullText);
    const truncated = text !== fullText;
    const meta = {k: kind, p: param.key};
    if (truncated) meta.tr = true;
    return {
      values: Object.assign({}, base, {param: param.key, value: text,
                                       valueType: param.type}),
      meta: meta
    };
  });
}

/**
 * Updates a parameter from an item row.
 * @param {!Object} draft
 * @param {!Object} row
 * @return {!Array<!Object>} Diffs.
 */
function updateParamItem_(draft, row) {
  const meta = row.sync.m || {};
  const v = row.valuesByKey;
  const original = findParam(draft.parameter, meta.p);
  if (!original) {
    throw new UserError(`The parameter "${meta.p}" was not found in GTM ` +
        'anymore (it may have been changed). List the sheet again.');
  }
  const key = toText(v.param);
  if (!key) throw new UserError('The parameter key is required.');
  const originalText = paramToText(original);
  if (meta.tr && toText(v.value) !== toText(cellSafeValue(originalText))) {
    throw new UserError('This value is too long to be edited in a cell. ' +
        'Edit it in Tag Manager.');
  }
  const diffs = [];
  let newParam = original;
  if (!meta.tr) {
    newParam = textToParam(key, v.value, original);
    if (paramToText(newParam) !== originalText ||
        newParam.type !== original.type) {
      diffs.push({header: key, from: originalText,
                  to: paramToText(newParam)});
    }
  }
  if (key !== meta.p) {
    if (findParam(draft.parameter, key)) {
      throw new UserError(`The parameter "${key}" already exists.`);
    }
    removeParam(draft, meta.p);
    diffs.push({header: 'Parameter key', from: meta.p, to: key});
  }
  newParam = Object.assign({}, newParam, {key: key});
  putParam(draft, newParam);
  row.newMeta = {k: meta.k, p: key};
  return diffs;
}

/**
 * Adds a parameter from an item row.
 * @param {!Object} draft
 * @param {!Object} row
 * @param {string} kind The entity kind.
 * @return {!Array<!Object>} Diffs.
 */
function createParamItem_(draft, row, kind) {
  const v = row.valuesByKey;
  const key = toText(v.param);
  if (!key) throw new UserError('The parameter key is required.');
  if (findParam(draft.parameter, key)) {
    throw new UserError(`The parameter "${key}" already exists in ` +
        `"${draft.name || draft.type}". Edit its row instead.`);
  }
  const param = textToParam(key, v.value);
  putParam(draft, param);
  row.newMeta = {k: kind || (row.sync.m && row.sync.m.k) || '', p: key};
  return [{header: key, from: '', to: paramToText(param) || '(empty)'}];
}

/**
 * Removes a parameter from an item row.
 * @param {!Object} draft
 * @param {!Object} row
 * @return {!Array<!Object>} Diffs.
 */
function deleteParamItem_(draft, row) {
  const meta = row.sync.m || {};
  const original = findParam(draft.parameter, meta.p);
  if (!original) {
    throw new UserError(`The parameter "${meta.p}" was not found in GTM ` +
        'anymore. List the sheet again.');
  }
  removeParam(draft, meta.p);
  return [{header: meta.p, from: truncateText(paramToText(original), 200),
           to: '(deleted)'}];
}
