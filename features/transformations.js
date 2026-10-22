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
 * @fileoverview The "Transformations" sheet (server containers).
 */

/** Transformation types. */
const TRANSFORMATION_TYPES = ['tf_allow_params', 'tf_exclude_params',
                              'tf_augment_event'];

/** Options of the "Affected tags" setting. */
const AFFECTED_TAGS_LABELS = {
  'true': 'All tags except',
  'false': 'Only these tags'
};

/**
 * Returns the definition of the Transformations sheet.
 * @return {!Object}
 */
function transformationsSheetDef_() {
  const rules = 'TRANSFORMATION RULES';
  const affected = 'AFFECTED TAGS';
  return {
    key: 'transformations',
    sheetName: 'Transformations',
    title: 'TRANSFORMATIONS (SERVER)',
    tabColor: TAB_COLORS.server,
    mode: 'entity',
    kind: 'transformation',
    contexts: ['server'],
    itemLabel: 'transformation',
    itemLabelPlural: 'transformations',
    idHeader: 'Transformation ID',
    nameColumn: 'name',
    columns: [
      COLUMNS.name('Transformation', {group: 'TRANSFORMATION'}),
      COLUMNS.type(TRANSFORMATION_TYPE_NAMES, 'Transformation',
          TRANSFORMATION_TYPES, {group: 'TRANSFORMATION'}),
      COLUMNS.param('priority', 'Priority', {
        group: 'TRANSFORMATION', width: 80,
        note: 'Whole number. Transformations with a higher priority run ' +
            'first.'
      }),
      COLUMNS.singleFieldTableParam('allowedParamsTable', 'Allowed Parameters',
          'allowedParams', {
            group: rules, types: ['tf_allow_params'],
            note: 'Comma-separated names of the event parameters that the ' +
                'affected tags can access.'
          }),
      COLUMNS.singleFieldTableParam('excludedParamsTable',
          'Excluded Parameters', 'excludedParams', {
            group: rules, types: ['tf_exclude_params'],
            note: 'Comma-separated names of the event parameters that are ' +
                'hidden from the affected tags.'
          }),
      COLUMNS.tableParam('augmentEventTable', 'Parameters to Add / Change',
          ['paramName', 'paramValue'], {
            group: rules, types: ['tf_augment_event'],
            note: 'One parameter per line: name = value.'
          }),
      affectedTagsModeColumn_({group: affected}),
      affectedTagsColumn_({group: affected}),
      matchingConditionsColumn_({group: 'MATCHING CONDITIONS'}),
      COLUMNS.folder({group: 'TRANSFORMATION SETTINGS'}),
      COLUMNS.notes({group: 'TRANSFORMATION SETTINGS'})
    ],
    newEntity: values => ({
      type: resolveTypeLabel(TRANSFORMATION_TYPE_NAMES, values.type),
      parameter: [
        {type: 'boolean', key: 'allTagsExcept', value: 'true'},
        {type: 'boolean', key: 'matchingConditionsEnabled', value: 'false'}
      ]
    }),
    validate: draft => {
      const priority = getParamValue(draft, 'priority');
      if (priority && !/^\d+$/.test(priority)) {
        return ['Priority must be a whole number (0 or more).'];
      }
      const table = {
        tf_allow_params: 'allowedParamsTable',
        tf_exclude_params: 'excludedParamsTable',
        tf_augment_event: 'augmentEventTable'
      }[draft.type];
      const param = table && findParam(draft.parameter, table);
      if (table && !(param && param.list && param.list.length)) {
        return ['Add at least one parameter to the transformation rules.'];
      }
      return [];
    }
  };
}

/**
 * Column: whether the transformation affects all tags except the listed ones,
 * or only the listed tags.
 * @param {!Object=} options
 * @return {!Object}
 */
function affectedTagsModeColumn_(options) {
  return buildColumn_({
    key: 'p:allTagsExcept',
    header: 'Affected Tags Mode',
    width: 150,
    dropdown: [AFFECTED_TAGS_LABELS['true'], AFFECTED_TAGS_LABELS['false']],
    strictDropdown: true,
    note: '"All tags except" the tags listed in "Affected Tags", or "Only ' +
        'these tags".',
    get: entity => AFFECTED_TAGS_LABELS[String(getBoolParam(entity,
        'allTagsExcept', true))],
    set: (draft, value) => {
      const text = toText(value).toLowerCase();
      const allExcept = text !== AFFECTED_TAGS_LABELS['false'].toLowerCase() &&
          text !== 'false' && text !== 'some tags';
      setBoolParam(draft, 'allTagsExcept', allExcept);
      const param = findParam(draft.parameter, 'affectedTags');
      (param && param.list || []).forEach(item => {
        (item.map || []).forEach(entry => {
          if (entry.type === 'tagReference') entry.isWeakReference = allExcept;
        });
      });
    }
  }, options);
}

/**
 * Column: tags affected by (or excluded from) a transformation.
 * @param {!Object=} options
 * @return {!Object}
 */
function affectedTagsColumn_(options) {
  return buildColumn_({
    key: 'p:affectedTags',
    header: 'Affected Tags',
    width: 240,
    kind: 'list',
    dropdown: 'Tags',
    note: 'Tag names, separated by commas (one per line if a name contains a ' +
        'comma).',
    get: entity => formatList(getTableRows(entity, 'affectedTags',
        ['tagReference']).map(row => row.name)),
    set: (draft, value, data) => {
      const names = parseList(value);
      const missing = names.filter(name => !data.getByName('tag', name));
      if (missing.length) {
        throw new UserError(`Tag not found: ${missing.join(', ')}.`);
      }
      if (!names.length) {
        removeParam(draft, 'affectedTags');
        return;
      }
      const weak = getBoolParam(draft, 'allTagsExcept', true);
      putParam(draft, {
        type: 'list',
        key: 'affectedTags',
        list: names.map(name => ({
          type: 'map',
          map: [{type: 'tagReference', key: 'tagReference', value: name,
                 isWeakReference: weak}]
        }))
      });
    }
  }, options);
}

/**
 * Column: matching conditions of a transformation (read-only summary).
 * @param {!Object=} options
 * @return {!Object}
 */
function matchingConditionsColumn_(options) {
  return buildColumn_({
    key: 'p:matchingConditionsTable',
    header: 'Matching Conditions',
    width: 260,
    kind: 'multiline',
    readOnly: true,
    note: 'Conditions under which the transformation applies (edit them in ' +
        'Tag Manager). Empty: the transformation always applies.',
    get: entity => {
      if (!getBoolParam(entity, 'matchingConditionsEnabled')) return '';
      const param = findParam(entity.parameter, 'matchingConditionsTable');
      return (param && param.list || []).map(item => [
        getMapField(item, 'variableReference') ||
            getMapField(item, 'variableName'),
        getMapField(item, 'expressionType'),
        getMapField(item, 'expressionValue')
      ].join(' ')).join('\n');
    }
  }, options);
}
