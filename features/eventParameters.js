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
 * @fileoverview The "Event Parameters & User Properties" sheet: one row per
 * event parameter, user property or configuration parameter of GA4 event
 * tags, Google tags, destination tags, Google Ads / Floodlight tags, Google
 * tag settings variables and server GA4 tags.
 */

/** Item types of the item-mode sheets. */
const ITEM_TYPE_LABELS = {
  tag: 'Tag',
  variable: 'Variable',
  client: 'Client',
  transformation: 'Transformation'
};

/**
 * Returns the parameter tables supported by the sheet.
 * @return {!Array<!Object>}
 */
function getParamTableDefs_() {
  const googleTags = ['googtag', 'dest_ga', 'dest_aw', 'dest_dc', 'dest_mc'];
  return [
    {setting: 'Event Parameter', key: 'eventSettingsTable',
     fields: ['parameter', 'parameterValue'],
     tagTypes: ['gaawe', 'awct', 'sp', 'awud', 'awcc', 'flc', 'fls'],
     variableTypes: ['gtes']},
    {setting: 'User Property', key: 'userProperties', fields: ['name', 'value'],
     tagTypes: ['gaawe'], variableTypes: ['gtes']},
    {setting: 'Configuration Parameter', key: 'configSettingsTable',
     fields: ['parameter', 'parameterValue'], tagTypes: googleTags,
     variableTypes: ['gtcs']},
    {setting: 'Shared Event Parameter', key: 'eventSettingsTable',
     fields: ['parameter', 'parameterValue'], tagTypes: googleTags,
     variableTypes: []},
    {setting: 'Parameter to Add / Edit', key: 'eventParameters',
     fields: ['fieldName', 'value'], tagTypes: ['sgtmgaaw'],
     variableTypes: []},
    {setting: 'Parameter to Include', key: 'epToInclude', fields: ['fieldName'],
     tagTypes: ['sgtmgaaw'], variableTypes: []},
    {setting: 'Parameter to Exclude', key: 'epToExclude', fields: ['fieldName'],
     tagTypes: ['sgtmgaaw'], variableTypes: []},
    {setting: 'User Property to Add / Edit', key: 'userProperties',
     fields: ['fieldName', 'value'], tagTypes: ['sgtmgaaw'],
     variableTypes: []},
    {setting: 'User Property to Include', key: 'upToInclude',
     fields: ['fieldName'], tagTypes: ['sgtmgaaw'], variableTypes: []},
    {setting: 'User Property to Exclude', key: 'upToExclude',
     fields: ['fieldName'], tagTypes: ['sgtmgaaw'], variableTypes: []}
  ];
}

/** Entity types that get an empty row when they have no parameters yet. */
const PARAM_PLACEHOLDER_TYPES = ['gaawe', 'gtes', 'sgtmgaaw'];

/**
 * Returns the parameter tables of an entity.
 * @param {!Object} entity
 * @param {string} kind 'tag' or 'variable'.
 * @return {!Array<!Object>}
 */
function getEntityParamTables_(entity, kind) {
  return getParamTableDefs_().filter(table =>
    (kind === 'tag' ? table.tagTypes : table.variableTypes)
        .indexOf(entity.type) !== -1);
}

/**
 * Returns the table of an entity that matches a setting name.
 * @param {!Object} entity
 * @param {string} kind
 * @param {*} setting
 * @return {!Object}
 */
function resolveParamTable_(entity, kind, setting) {
  const tables = getEntityParamTables_(entity, kind);
  const text = toText(setting).toLowerCase();
  if (!text && tables.length) return tables[0];
  const table = tables.find(t => t.setting.toLowerCase() === text);
  if (!table) {
    throw new UserError(`"${toText(setting)}" can't be used with ` +
        `"${entity.name}". Use one of: ` +
        tables.map(t => t.setting).join(', ') + '.');
  }
  return table;
}

/**
 * Returns the rows (map items) of a table parameter.
 * @param {!Object} entity
 * @param {string} key
 * @return {!Array<!Object>}
 */
function getTableItems_(entity, key) {
  const param = findParam(entity.parameter, key);
  return param && param.list ? param.list : [];
}

/**
 * Replaces the rows of a table parameter (removes it when empty).
 * @param {!Object} entity
 * @param {string} key
 * @param {!Array<!Object>} items
 */
function setTableItems_(entity, key, items) {
  if (!items.length) {
    removeParam(entity, key);
  } else {
    putParam(entity, {type: 'list', key: key, list: items});
  }
}

/**
 * Sets a field of a map item.
 * @param {!Object} item
 * @param {string} field
 * @param {string} value
 */
function setMapField_(item, field, value) {
  item.map = item.map || [];
  const entry = item.map.find(e => e.key === field);
  if (entry) {
    entry.value = value;
  } else {
    item.map.push({type: 'template', key: field, value: value});
  }
}

/**
 * Returns the definition of the Event Parameters & User Properties sheet.
 * @return {!Object}
 */
function eventParametersSheetDef_() {
  const settings = [];
  getParamTableDefs_().forEach(table => {
    if (settings.indexOf(table.setting) === -1) settings.push(table.setting);
  });
  return {
    key: 'eventParameters',
    sheetName: 'Event Parameters & User Properties',
    title: 'EVENT PARAMETERS & USER PROPERTIES',
    tabColor: TAB_COLORS.general,
    mode: 'item',
    kinds: ['tag', 'variable'],
    contexts: ['web', 'server'],
    itemLabel: 'parameter',
    itemLabelPlural: 'parameters and properties',
    idHeader: 'Item ID',
    nameColumn: 'itemName',
    columns: [
      {key: 'itemType', header: 'Item Type', group: 'TAG OR VARIABLE',
       width: 80, createOnly: true, dropdown: ['Tag', 'Variable'],
       note: 'Tag or Variable.'},
      {key: 'itemName', header: 'Tag / Variable Name', group: 'TAG OR VARIABLE',
       width: 280, createOnly: true, dropdown: 'Tags',
       note: 'Name of the tag or variable. For new rows, the name is used to ' +
           'find the tag or variable.'},
      {key: 'itemKind', header: 'Tag / Variable Type', group: 'TAG OR VARIABLE',
       width: 190, readOnly: true},
      {key: 'eventName', header: 'Event Name', group: 'TAG OR VARIABLE',
       width: 150, readOnly: true,
       note: 'Event name of the tag (for information and filtering).'},
      {key: 'setting', header: 'Setting', group: 'PARAMETER', width: 190,
       dropdown: settings, strictDropdown: true,
       note: 'Event Parameter, User Property, Configuration Parameter, ' +
           'Shared Event Parameter (Google tags) or one of the server GA4 ' +
           'tag settings.'},
      {key: 'name', header: 'Parameter / Property Name', group: 'PARAMETER',
       width: 220},
      {key: 'value', header: 'Value', group: 'PARAMETER', width: 260,
       dropdown: 'Variables',
       note: 'Value of the parameter, e.g. {{Page Type}} (not used for ' +
           '"to Include" / "to Exclude" settings).'}
    ],
    isEligible: (entity, kind) => getEntityParamTables_(entity, kind).length > 0,
    explode: (entity, kind, data, ctx) => {
      const base = {
        itemType: ITEM_TYPE_LABELS[kind],
        itemName: entity.name,
        itemKind: getTypeLabel(kind === 'tag' ? TAG_TYPE_NAMES :
            VARIABLE_TYPE_NAMES, entity.type, data),
        eventName: entity.type === 'gaawe' || entity.type === 'sgtmgaaw' ?
            getParamValue(entity, 'eventName') : ''
      };
      const rows = [];
      const tables = getEntityParamTables_(entity, kind);
      tables.forEach(table => {
        getTableRows(entity, table.key, table.fields).forEach(row => {
          rows.push({
            values: Object.assign({}, base, {
              setting: table.setting, name: row.name, value: row.value}),
            meta: {k: kind, s: table.setting, n: row.name, v: row.value}
          });
        });
      });
      if (!rows.length && PARAM_PLACEHOLDER_TYPES.indexOf(entity.type) !== -1) {
        rows.push({
          values: Object.assign({}, base, {setting: tables[0].setting,
                                           name: '', value: ''}),
          meta: {k: kind, placeholder: true}
        });
      }
      return rows;
    },
    findEntity: (row, data, ctx, op) => findItemEntity_(row, data,
        ['tag', 'variable'],
        (entity, kind) => getEntityParamTables_(entity, kind).length > 0, op),
    updateItem: (draft, row, data, ctx) => {
      const meta = row.sync.m || {};
      if (meta.placeholder) return createParamTableItem_(draft, row, meta.k);
      checkItemNameUnchanged_(draft, row);
      const v = row.valuesByKey;
      const oldTable = resolveParamTable_(draft, meta.k, meta.s);
      const newTable = resolveParamTable_(draft, meta.k, v.setting);
      const name = toText(v.name);
      const value = newTable.fields[1] ? toText(v.value) : '';
      if (!name) throw new UserError('The name is required.');
      const items = getTableItems_(draft, oldTable.key);
      const index = items.findIndex(item =>
        getMapField(item, oldTable.fields[0]) === meta.n &&
        (!oldTable.fields[1] || getMapField(item, oldTable.fields[1]) ===
            meta.v));
      if (index === -1) {
        throw new UserError(`${oldTable.setting} "${meta.n}" was not found ` +
            'in GTM anymore (it may have been changed). List the sheet again.');
      }
      const diffs = [];
      if (oldTable.setting === newTable.setting) {
        const item = items[index];
        if (name !== meta.n) {
          setMapField_(item, oldTable.fields[0], name);
          diffs.push({header: `${oldTable.setting} name`, from: meta.n,
                      to: name});
        }
        if (oldTable.fields[1] && value !== toText(meta.v)) {
          setMapField_(item, oldTable.fields[1], value);
          diffs.push({header: `${oldTable.setting} "${name}"`,
                      from: toText(meta.v), to: value});
        }
        setTableItems_(draft, oldTable.key, items);
      } else {
        items.splice(index, 1);
        setTableItems_(draft, oldTable.key, items);
        const target = getTableItems_(draft, newTable.key);
        target.push(buildTableItem(newTable.fields, {name: name,
                                                      value: value}));
        setTableItems_(draft, newTable.key, target);
        diffs.push({header: `${oldTable.setting} "${meta.n}"`,
                    from: toText(meta.v), to: `(moved to ${newTable.setting})`},
                   {header: `${newTable.setting} "${name}"`, from: '',
                    to: value});
      }
      row.newMeta = {k: meta.k, s: newTable.setting, n: name, v: value};
      return diffs;
    },
    createItem: (draft, row, data, ctx, kind) =>
      createParamTableItem_(draft, row, kind),
    deleteItem: (draft, row) => {
      const meta = row.sync.m || {};
      if (meta.placeholder) {
        throw new UserError('This row is an empty placeholder; there is ' +
            'nothing to delete.');
      }
      const table = resolveParamTable_(draft, meta.k, meta.s);
      const items = getTableItems_(draft, table.key);
      const index = items.findIndex(item =>
        getMapField(item, table.fields[0]) === meta.n &&
        (!table.fields[1] || getMapField(item, table.fields[1]) === meta.v));
      if (index === -1) {
        throw new UserError(`${table.setting} "${meta.n}" was not found in ` +
            'GTM anymore. List the sheet again.');
      }
      items.splice(index, 1);
      setTableItems_(draft, table.key, items);
      return [{header: `${table.setting} "${meta.n}"`, from: toText(meta.v),
               to: '(deleted)'}];
    }
  };
}

/**
 * Adds a row to a parameter table of an entity.
 * @param {!Object} draft
 * @param {!Object} row
 * @param {string} kind
 * @return {!Array<!Object>} Diffs.
 */
function createParamTableItem_(draft, row, kind) {
  const v = row.valuesByKey;
  const table = resolveParamTable_(draft, kind, v.setting);
  const name = toText(v.name);
  const value = table.fields[1] ? toText(v.value) : '';
  if (!name) throw new UserError('The name is required.');
  const items = getTableItems_(draft, table.key);
  if (items.some(item => getMapField(item, table.fields[0]) === name)) {
    throw new UserError(`${table.setting} "${name}" already exists in ` +
        `"${draft.name}". Edit its row instead.`);
  }
  items.push(buildTableItem(table.fields, {name: name, value: value}));
  setTableItems_(draft, table.key, items);
  row.newMeta = {k: kind, s: table.setting, n: name, v: value};
  return [{header: `${table.setting} "${name}"`, from: '',
           to: value || '(added)'}];
}

/**
 * Finds the entity of an item row (by ID, or by type and name for new rows).
 * @param {!Object} row
 * @param {!WorkspaceData} data
 * @param {!Array<string>} kinds Allowed entity kinds.
 * @param {function(!Object, string): boolean} isEligible
 * @param {!Object=} op The operation. New rows ("Create") are matched by name,
 *     so that a copied row can add an item to another tag or variable.
 * @return {{kind: string, entity: !Object}}
 */
function findItemEntity_(row, data, kinds, isEligible, op) {
  const v = row.valuesByKey;
  const meta = row.sync.m;
  const typeText = toText(v.itemType).toLowerCase();
  const byName = Boolean(op && op.type === 'create' && toText(v.itemName));
  const id = byName ? '' : row.id;
  let candidates = kinds.filter(kind => workspaceSupportsKind(data.ctx, kind));
  if (meta && meta.k && !byName) {
    candidates = [meta.k];
  } else if (typeText) {
    candidates = candidates.filter(kind =>
      ITEM_TYPE_LABELS[kind].toLowerCase() === typeText || kind === typeText);
    if (!candidates.length) {
      throw new UserError(`Unknown item type "${toText(v.itemType)}". Use ` +
          kinds.map(kind => ITEM_TYPE_LABELS[kind]).join(', ') + '.');
    }
  }
  const name = toText(v.itemName);
  for (const kind of candidates) {
    const entity = id ? data.getById(kind, id) :
        (name ? data.getByName(kind, name) : null);
    if (!entity) continue;
    if (!isEligible(entity, kind)) {
      throw new UserError(`"${entity.name}" can't be used in this sheet ` +
          '(this setting is not available for its type).');
    }
    return {kind: kind, entity: entity};
  }
  if (id) {
    throw new UserError(`The item with ID ${id} was not found (it may ` +
        'have been deleted). List the sheet again.');
  }
  throw new UserError(name ? `"${name}" was not found. Check the name and ` +
      'the item type.' : 'Enter the name of the tag or variable.');
}

/**
 * Throws if the item name of a listed row was changed.
 * @param {!Object} draft
 * @param {!Object} row
 */
function checkItemNameUnchanged_(draft, row) {
  const name = toText(row.valuesByKey.itemName);
  if (name && name !== draft.name) {
    throw new UserError('The name of the tag or variable can\'t be changed ' +
        'in this sheet. Rename it in its own sheet (e.g. Tags), then list ' +
        'this sheet again.');
  }
}
