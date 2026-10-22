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
 * @fileoverview The "Variables" sheet: all user-defined variables, with their
 * main setting and where they are used (web and server containers).
 */

/** Parameter that holds the main value of each variable type. */
const VARIABLE_VALUE_PARAMS = {
  v: 'name',
  c: 'value',
  jsm: 'javascript',
  k: 'name',
  j: 'name',
  aev: 'varType',
  smm: 'input',
  remm: 'input',
  awec: 'mode',
  ed: 'keyPath',
  rh: 'headerName',
  qp: 'queryParamName',
  sgtmk: 'name'
};

/** Variable types that support a default value. */
const VARIABLE_DEFAULT_VALUE_TYPES = ['v', 'aev', 'smm', 'remm', 'ed'];

/** Variable types that can be created in the Variables sheet. */
const CREATABLE_VARIABLE_TYPES = ['v', 'c', 'jsm', 'k', 'j', 'u', 'f', 'aev',
                                  'd', 'smm', 'remm', 'gtes', 'gtcs', 'uv',
                                  'ed', 'rh', 'qp', 'sgtmk'];

/** Notes explaining the main value per variable type. */
const VARIABLE_VALUE_NOTE = 'Main setting of the variable:\n' +
    '• Data Layer Variable: data layer key\n' +
    '• Constant: value\n' +
    '• Custom JavaScript: the function code\n' +
    '• 1st-Party Cookie / Cookie Value: cookie name\n' +
    '• JavaScript Variable: global variable name\n' +
    '• URL / HTTP Referrer: component (URL, PROTOCOL, HOST, PORT, PATH, ' +
    'EXTENSION, FRAGMENT, or "QUERY: key")\n' +
    '• DOM Element / Element Visibility: "ID: element-id" or "CSS: selector"\n' +
    '• Auto-Event Variable: variable type (ELEMENT, CLASSES, ID, TARGET, ' +
    'TEXT, URL, ...)\n' +
    '• Lookup / RegEx Table: input variable\n' +
    '• Event Data: key path\n' +
    '• Request Header: header name\n' +
    '• Query Parameter: parameter name\n' +
    'Other settings can be edited in the "Advanced Settings" sheet.';

/**
 * Returns the definition of the Variables sheet.
 * @return {!Object}
 */
function variablesSheetDef_() {
  return {
    key: 'variables',
    sheetName: 'Variables',
    title: 'VARIABLES',
    tabColor: TAB_COLORS.general,
    mode: 'entity',
    kind: 'variable',
    contexts: ['web', 'server'],
    itemLabel: 'variable',
    itemLabelPlural: 'variables',
    idHeader: 'Variable ID',
    nameColumn: 'name',
    columns: [
      COLUMNS.name('Variable', {group: 'VARIABLE'}),
      COLUMNS.type(VARIABLE_TYPE_NAMES, 'Variable', CREATABLE_VARIABLE_TYPES,
          {group: 'VARIABLE'}),
      variableValueColumn_({group: 'VARIABLE'}),
      variableDefaultValueColumn_({group: 'VARIABLE'}),
      COLUMNS.tableParam('map', 'Lookup Table (input = output)',
          ['key', 'value'], {
            group: 'VARIABLE', types: ['smm', 'remm'],
            note: 'One row per line: input = output, e.g.\n' +
                'www.example.com = G-AAAAAAAAAA\n' +
                'staging.example.com = G-BBBBBBBBBB'
          }),
      variableFormatColumn_({group: 'VARIABLE'}),
      COLUMNS.folder({group: 'VARIABLE'}),
      COLUMNS.notes({group: 'VARIABLE'}),
      variableUsageCountColumn_({group: 'USAGE'}),
      variableUsageColumn_({group: 'USAGE'})
    ],
    newEntity: values => {
      const type = resolveTypeLabel(VARIABLE_TYPE_NAMES, values.type);
      return {type: type, parameter: defaultVariableParameters_(type)};
    },
    validateDelete: (variable, data, ctx) => {
      const usage = getVariableUsage_(variable.name, data, ctx);
      if (!usage.length) return [];
      return [`This variable can't be deleted because it is used in ` +
          `${usage.length} item(s): ${formatUsage_(usage)}. Remove these ` +
          'references first.'];
    }
  };
}

/**
 * Returns the default parameters of a new variable.
 * @param {string} type
 * @return {!Array<!Object>}
 */
function defaultVariableParameters_(type) {
  const bool = (key, value) => ({type: 'boolean', key: key,
                                 value: value ? 'true' : 'false'});
  switch (type) {
    case 'v':
      return [{type: 'integer', key: 'dataLayerVersion', value: '2'},
              bool('setDefaultValue', false)];
    case 'k':
    case 'sgtmk':
      return [bool('decodeCookie', false)];
    case 'u':
    case 'f':
      return [{type: 'template', key: 'component', value: 'URL'}];
    case 'aev':
      return [{type: 'template', key: 'varType', value: 'ELEMENT'},
              bool('setDefaultValue', false)];
    case 'd':
      return [{type: 'template', key: 'selectorType', value: 'ID'}];
    case 'smm':
      return [bool('setDefaultValue', false)];
    case 'remm':
      return [bool('setDefaultValue', false), bool('ignoreCase', true),
              bool('fullMatch', true), bool('replaceAfterMatch', true)];
    case 'ed':
      return [bool('setDefaultValue', false)];
    default:
      return [];
  }
}

/**
 * Column: main setting of a variable.
 * @param {!Object=} options
 * @return {!Object}
 */
function variableValueColumn_(options) {
  const isSelectorType = type => type === 'd' || type === 'vis';
  const isUrlType = type => type === 'u' || type === 'f';
  return buildColumn_({
    key: 'value',
    header: 'Value / Main Setting',
    width: 260,
    kind: 'multiline',
    note: VARIABLE_VALUE_NOTE,
    appliesTo: entity => Boolean(VARIABLE_VALUE_PARAMS[entity.type]) ||
        isSelectorType(entity.type) || isUrlType(entity.type),
    get: entity => {
      if (isSelectorType(entity.type)) {
        return getParamValue(entity, 'selectorType') === 'CSS' ?
            'CSS: ' + getParamValue(entity, 'elementSelector') :
            'ID: ' + getParamValue(entity, 'elementId');
      }
      if (isUrlType(entity.type)) {
        const component = getParamValue(entity, 'component') || 'URL';
        const queryKey = getParamValue(entity, 'queryKey');
        return component === 'QUERY' && queryKey ? `QUERY: ${queryKey}` :
            component;
      }
      return getParamValue(entity, VARIABLE_VALUE_PARAMS[entity.type]);
    },
    set: (draft, value) => {
      const text = toText(value);
      if (isSelectorType(draft.type)) {
        const match = /^(css|id)\s*:\s*(.*)$/i.exec(text);
        const isCss = match ? match[1].toLowerCase() === 'css' : false;
        const selector = match ? match[2].trim() : text;
        setParamValue(draft, 'selectorType', isCss ? 'CSS' : 'ID');
        setParamValue(draft, isCss ? 'elementSelector' : 'elementId',
            selector);
        removeParam(draft, isCss ? 'elementId' : 'elementSelector');
        return;
      }
      if (isUrlType(draft.type)) {
        const match = /^query\s*:\s*(.+)$/i.exec(text);
        if (match) {
          setParamValue(draft, 'component', 'QUERY');
          setParamValue(draft, 'queryKey', match[1].trim());
        } else {
          setParamValue(draft, 'component', (text || 'URL').toUpperCase());
          removeParam(draft, 'queryKey');
        }
        return;
      }
      setParamValue(draft, VARIABLE_VALUE_PARAMS[draft.type], text);
    }
  }, options);
}

/**
 * Column: default value of a variable (empty = no default value).
 * @param {!Object=} options
 * @return {!Object}
 */
function variableDefaultValueColumn_(options) {
  return buildColumn_({
    key: 'defaultValue',
    header: 'Default Value',
    width: 140,
    note: 'Value returned when the variable is undefined (Data Layer, ' +
        'Auto-Event, Lookup Table, RegEx Table and Event Data variables). ' +
        'Leave empty for no default value.',
    appliesTo: entity => VARIABLE_DEFAULT_VALUE_TYPES.indexOf(entity.type) !==
        -1,
    get: entity => getBoolParam(entity, 'setDefaultValue') ?
        getParamValue(entity, 'defaultValue') : '',
    set: (draft, value) => {
      const text = toText(value);
      setBoolParam(draft, 'setDefaultValue', Boolean(text));
      if (text) {
        setParamValue(draft, 'defaultValue', text);
      } else {
        removeParam(draft, 'defaultValue');
      }
    }
  }, options);
}

/** Case conversion options of the variable format settings. */
const CASE_CONVERSION_LABELS = {none: 'None', lowercase: 'Lowercase',
                                uppercase: 'Uppercase'};

/**
 * Column: case conversion of the variable value (Format Value).
 * @param {!Object=} options
 * @return {!Object}
 */
function variableFormatColumn_(options) {
  return buildColumn_({
    key: 'caseConversion',
    header: 'Convert Case (Format Value)',
    width: 120,
    dropdown: ['None', 'Lowercase', 'Uppercase'],
    strictDropdown: true,
    get: entity => {
      const type = entity.formatValue && entity.formatValue.caseConversionType;
      return CASE_CONVERSION_LABELS[type || 'none'] || 'None';
    },
    set: (draft, value) => {
      const text = toText(value).toLowerCase() || 'none';
      const type = Object.keys(CASE_CONVERSION_LABELS).find(key =>
        CASE_CONVERSION_LABELS[key].toLowerCase() === text || key === text);
      if (!type) throw new UserError(`Unknown case conversion "${value}".`);
      draft.formatValue = draft.formatValue || {};
      if (type === 'none') {
        delete draft.formatValue.caseConversionType;
      } else {
        draft.formatValue.caseConversionType = type;
      }
      if (!Object.keys(draft.formatValue).length) delete draft.formatValue;
    }
  }, options);
}

/**
 * Collects the names of the variables referenced by an entity.
 * @param {*} value An entity or a part of it.
 * @param {!Object<string, boolean>} names Collected names.
 */
function collectVariableReferences_(value, names) {
  if (typeof value === 'string') {
    const pattern = /\{\{([^{}]+)\}\}/g;
    let match;
    while ((match = pattern.exec(value)) !== null) names[match[1]] = true;
  } else if (Array.isArray(value)) {
    value.forEach(item => collectVariableReferences_(item, names));
  } else if (value && typeof value === 'object') {
    Object.keys(value).forEach(key => {
      if (key !== 'name' && key !== 'notes' && key !== 'path' &&
          key !== 'tagManagerUrl') {
        collectVariableReferences_(value[key], names);
      }
    });
  }
}

/**
 * Returns an index of variable name to the items that reference it.
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Object<string, !Array<{kind: string, name: string}>>}
 */
function getVariableUsageIndex_(data, ctx) {
  if (data.variableUsageIndex_) return data.variableUsageIndex_;
  const index = {};
  ['tag', 'trigger', 'variable', 'client', 'transformation'].forEach(kind => {
    if (!workspaceSupportsKind(ctx, kind)) return;
    data.list(kind).forEach(entity => {
      const names = {};
      collectVariableReferences_(entity, names);
      Object.keys(names).forEach(name => {
        (index[name] = index[name] || []).push(
            {kind: kind, name: entity.name});
      });
    });
  });
  data.variableUsageIndex_ = index;
  return index;
}

/**
 * Returns the items that use a variable.
 * @param {string} variableName
 * @param {!WorkspaceData} data
 * @param {!Object} ctx
 * @return {!Array<{kind: string, name: string}>}
 */
function getVariableUsage_(variableName, data, ctx) {
  return (getVariableUsageIndex_(data, ctx)[variableName] || [])
      .filter(item => !(item.kind === 'variable' &&
                        item.name === variableName));
}

/**
 * Formats a usage list, e.g. "Tags: A, B; Triggers: C".
 * @param {!Array<{kind: string, name: string}>} usage
 * @param {string=} separator
 * @return {string}
 */
function formatUsage_(usage, separator) {
  const labels = {tag: 'Tags', trigger: 'Triggers', variable: 'Variables',
                  client: 'Clients', transformation: 'Transformations'};
  return Object.keys(labels)
      .map(kind => {
        const names = usage.filter(item => item.kind === kind)
            .map(item => item.name);
        return names.length ? `${labels[kind]}: ${names.join(', ')}` : '';
      })
      .filter(Boolean)
      .join(separator || '; ');
}

/**
 * Column: number of items that use the variable (read-only).
 * @param {!Object=} options
 * @return {!Object}
 */
function variableUsageCountColumn_(options) {
  return buildColumn_({
    key: 'usageCount',
    header: 'Used In (count)',
    width: 90,
    readOnly: true,
    note: 'Number of tags, triggers, variables, clients and transformations ' +
        'that reference this variable. 0 = unused.',
    get: (entity, data, ctx) => String(getVariableUsage_(entity.name, data,
        ctx).length)
  }, options);
}

/**
 * Column: items that use the variable (read-only).
 * @param {!Object=} options
 * @return {!Object}
 */
function variableUsageColumn_(options) {
  return buildColumn_({
    key: 'usage',
    header: 'Used In',
    width: 320,
    kind: 'multiline',
    readOnly: true,
    get: (entity, data, ctx) => formatUsage_(getVariableUsage_(entity.name,
        data, ctx), '\n')
  }, options);
}
