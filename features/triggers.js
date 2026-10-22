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
 * @fileoverview The "Triggers" sheet (web and server containers). Trigger
 * conditions are written as readable text, one condition per line, e.g.
 * "{{Page Path}} contains /checkout".
 */

/** Trigger types that can be created in the Triggers sheet. */
const CREATABLE_TRIGGER_TYPES = ['customEvent', 'pageview', 'domReady',
                                 'windowLoaded', 'click', 'linkClick',
                                 'formSubmission', 'historyChange', 'jsError',
                                 'init', 'consentInit', 'always',
                                 'serverPageview'];

/** Trigger types with "Enable this trigger when" conditions. */
const AUTO_EVENT_FILTER_TRIGGER_TYPES = ['linkClick', 'formSubmission',
                                         'timer'];

/** Condition operators, as shown in the sheet. */
const CONDITION_OPERATORS = [
  {label: 'equals', type: 'equals'},
  {label: 'does not equal', type: 'equals', negate: true},
  {label: 'contains', type: 'contains'},
  {label: 'does not contain', type: 'contains', negate: true},
  {label: 'starts with', type: 'startsWith'},
  {label: 'does not start with', type: 'startsWith', negate: true},
  {label: 'ends with', type: 'endsWith'},
  {label: 'does not end with', type: 'endsWith', negate: true},
  {label: 'matches RegEx', type: 'matchRegex'},
  {label: 'does not match RegEx', type: 'matchRegex', negate: true},
  {label: 'matches CSS selector', type: 'cssSelector'},
  {label: 'does not match CSS selector', type: 'cssSelector', negate: true},
  {label: 'less than', type: 'less'},
  {label: 'less than or equal to', type: 'lessOrEquals'},
  {label: 'greater than', type: 'greater'},
  {label: 'greater than or equal to', type: 'greaterOrEquals'},
  {label: 'matches URL', type: 'urlMatches'}
];

/** Note explaining the condition format. */
const CONDITIONS_NOTE = 'One condition per line (all conditions must be ' +
    'true): {{Variable}} operator value, e.g.\n' +
    '{{Page Path}} contains /checkout\n' +
    '{{Click Classes}} does not contain button\n' +
    '{{Page URL}} matches RegEx (ignore case) ^https://shop\\.\n' +
    'Operators: ' + CONDITION_OPERATORS.map(op => op.label).join(', ') +
    '. Add "(ignore case)" after the operator to ignore case.';

/**
 * Returns the definition of the Triggers sheet.
 * @return {!Object}
 */
function triggersSheetDef_() {
  return {
    key: 'triggers',
    sheetName: 'Triggers',
    title: 'TRIGGERS',
    tabColor: TAB_COLORS.general,
    mode: 'entity',
    kind: 'trigger',
    contexts: ['web', 'server'],
    itemLabel: 'trigger',
    itemLabelPlural: 'triggers',
    idHeader: 'Trigger ID',
    nameColumn: 'name',
    columns: [
      COLUMNS.name('Trigger', {group: 'TRIGGER'}),
      COLUMNS.type(TRIGGER_TYPE_NAMES, 'Trigger', CREATABLE_TRIGGER_TYPES, {
        group: 'TRIGGER',
        note: 'Trigger type. Other types (e.g. Timer, Scroll Depth, Element ' +
            'Visibility) can be created by copying an existing trigger row.'
      }),
      customEventNameColumn_({group: 'TRIGGER'}),
      customEventRegexColumn_({group: 'TRIGGER'}),
      conditionsColumn_('filter', 'Conditions (fires on)', {
        group: 'TRIGGER', note: CONDITIONS_NOTE +
            '\nLeave empty to fire on all events of this type.'
      }),
      conditionsColumn_('autoEventFilter', 'Enable When', {
        group: 'TRIGGER', types: AUTO_EVENT_FILTER_TRIGGER_TYPES,
        note: 'Click (just links), form and timer triggers: the trigger only ' +
            'listens when these conditions are true. ' + CONDITIONS_NOTE
      }),
      COLUMNS.folder({group: 'TRIGGER'}),
      COLUMNS.notes({group: 'TRIGGER'}),
      triggerUsageColumn_({group: 'USAGE'})
    ],
    newEntity: values => {
      const type = resolveTypeLabel(TRIGGER_TYPE_NAMES, values.type);
      const entity = {type: type};
      if (type === 'linkClick' || type === 'formSubmission') {
        entity.waitForTags = {type: 'boolean', value: 'false'};
        entity.checkValidation = {type: 'boolean', value: 'false'};
      }
      return entity;
    },
    validate: draft => {
      if (draft.type === 'customEvent' &&
          !(draft.customEventFilter && draft.customEventFilter.length)) {
        return ['Event Name is required for Custom Event triggers.'];
      }
      return [];
    },
    validateDelete: (trigger, data) => {
      const usage = getTriggerUsage_(trigger.triggerId, data);
      if (!usage.length) return [];
      return [`This trigger can't be deleted because it is used by: ` +
          `${usage.join(', ')}. Remove it from these tags first.`];
    }
  };
}

/**
 * Formats a trigger condition.
 * @param {!Object} condition A Condition resource.
 * @return {string}
 */
function formatCondition_(condition) {
  const params = condition.parameter || [];
  const value = key => {
    const param = findParam(params, key);
    return param && param.value !== undefined ? String(param.value) : '';
  };
  const negate = value('negate') === 'true';
  const ignoreCase = value('ignore_case') === 'true';
  const operator = CONDITION_OPERATORS.find(op => op.type === condition.type &&
      Boolean(op.negate) === negate);
  let label = operator ? operator.label :
      condition.type + (negate ? ' (negated)' : '');
  if (ignoreCase) label += ' (ignore case)';
  return `${value('arg0')} ${label} ${value('arg1')}`.trim();
}

/**
 * Parses a trigger condition.
 * @param {string} line E.g. "{{Page Path}} contains /checkout".
 * @return {!Object} A Condition resource.
 */
function parseCondition_(line) {
  const match = /^(\{\{[^{}]+\}\})\s*(.*)$/.exec(line.trim());
  if (!match) {
    throw new UserError(`The condition "${line}" must start with a ` +
        'variable, e.g. {{Page Path}} contains /checkout');
  }
  const rest = match[2];
  const lower = rest.toLowerCase();
  const operators = CONDITION_OPERATORS.slice()
      .sort((a, b) => b.label.length - a.label.length);
  for (const operator of operators) {
    const label = operator.label.toLowerCase();
    if (lower.indexOf(label) !== 0 ||
        (rest.length > label.length && rest.charAt(label.length) !== ' ')) {
      continue;
    }
    let after = rest.slice(label.length);
    let ignoreCase = false;
    const ignoreMatch = /^\s*\(ignore case\)/i.exec(after);
    if (ignoreMatch) {
      ignoreCase = true;
      after = after.slice(ignoreMatch[0].length);
    }
    const parameter = [
      {type: 'template', key: 'arg0', value: match[1]},
      {type: 'template', key: 'arg1', value: after.trim()}
    ];
    if (ignoreCase) {
      parameter.push({type: 'boolean', key: 'ignore_case', value: 'true'});
    }
    if (operator.negate) {
      parameter.push({type: 'boolean', key: 'negate', value: 'true'});
    }
    return {type: operator.type, parameter: parameter};
  }
  throw new UserError(`Unknown operator in the condition "${line}". Use ` +
      'one of: ' + CONDITION_OPERATORS.map(op => op.label).join(', ') + '.');
}

/**
 * Column: a list of trigger conditions.
 * @param {string} field 'filter' or 'autoEventFilter'.
 * @param {string} header
 * @param {!Object=} options
 * @return {!Object}
 */
function conditionsColumn_(field, header, options) {
  return buildColumn_({
    key: field,
    header: header,
    width: 300,
    kind: 'multiline',
    get: entity => (entity[field] || []).map(formatCondition_).join('\n'),
    set: (draft, value) => {
      const lines = toText(value).split('\n').map(line => line.trim())
          .filter(Boolean);
      if (lines.length) {
        draft[field] = lines.map(parseCondition_);
      } else {
        delete draft[field];
      }
    }
  }, options);
}

/**
 * Returns the event name condition of a Custom Event trigger, if it has the
 * usual form ({{_event}} equals / matches RegEx a value).
 * @param {!Object} trigger
 * @return {?{value: string, regex: boolean}}
 */
function getCustomEventName_(trigger) {
  const filters = trigger.customEventFilter || [];
  if (filters.length !== 1) return null;
  const condition = filters[0];
  if (condition.type !== 'equals' && condition.type !== 'matchRegex') {
    return null;
  }
  const arg0 = findParam(condition.parameter, 'arg0');
  const arg1 = findParam(condition.parameter, 'arg1');
  if (!arg0 || arg0.value !== '{{_event}}') return null;
  return {value: arg1 ? String(arg1.value || '') : '',
          regex: condition.type === 'matchRegex'};
}

/**
 * Builds the event name condition of a Custom Event trigger.
 * @param {string} name
 * @param {boolean} regex
 * @param {boolean} ignoreCase
 * @return {!Array<!Object>}
 */
function buildCustomEventFilter_(name, regex, ignoreCase) {
  const parameter = [
    {type: 'template', key: 'arg0', value: '{{_event}}'},
    {type: 'template', key: 'arg1', value: name}
  ];
  if (regex && ignoreCase) {
    parameter.push({type: 'boolean', key: 'ignore_case', value: 'true'});
  }
  return [{type: regex ? 'matchRegex' : 'equals', parameter: parameter}];
}

/**
 * Column: event name of a Custom Event trigger.
 * @param {!Object=} options
 * @return {!Object}
 */
function customEventNameColumn_(options) {
  return buildColumn_({
    key: 'customEventName',
    header: 'Event Name (Custom Event)',
    width: 180,
    note: 'The dataLayer event name that fires a Custom Event trigger, ' +
        'e.g. purchase.',
    appliesTo: entity => entity.type === 'customEvent',
    get: entity => {
      const event = getCustomEventName_(entity);
      if (event) return event.value;
      return (entity.customEventFilter || []).map(formatCondition_)
          .join('\n');
    },
    set: (draft, value) => {
      const name = toText(value);
      if (!name) {
        delete draft.customEventFilter;
        return;
      }
      const current = getCustomEventName_(draft);
      const condition = (draft.customEventFilter || [])[0];
      const ignoreCase = Boolean(condition && findParam(condition.parameter,
          'ignore_case'));
      draft.customEventFilter = buildCustomEventFilter_(name,
          Boolean(current && current.regex), ignoreCase);
    }
  }, options);
}

/**
 * Column: whether the event name of a Custom Event trigger is a regular
 * expression.
 * @param {!Object=} options
 * @return {!Object}
 */
function customEventRegexColumn_(options) {
  return buildColumn_({
    key: 'customEventRegex',
    header: 'Use RegEx Matching',
    width: 100,
    kind: 'checkbox',
    appliesTo: entity => entity.type === 'customEvent',
    get: entity => {
      const event = getCustomEventName_(entity);
      return Boolean(event && event.regex);
    },
    set: (draft, value) => {
      const event = getCustomEventName_(draft);
      if (!event) {
        if (toBool(value)) {
          throw new UserError('Enter an Event Name to use RegEx matching.');
        }
        return;
      }
      draft.customEventFilter = buildCustomEventFilter_(event.value,
          toBool(value), false);
    }
  }, options);
}

/**
 * Returns the names of the tags and trigger groups that use a trigger.
 * @param {string} triggerId
 * @param {!WorkspaceData} data
 * @return {!Array<string>}
 */
function getTriggerUsage_(triggerId, data) {
  const id = String(triggerId);
  const usage = [];
  data.list('tag').forEach(tag => {
    const firing = (tag.firingTriggerId || []).map(String);
    const blocking = (tag.blockingTriggerId || []).map(String);
    if (firing.indexOf(id) !== -1) usage.push(`${tag.name} (firing)`);
    if (blocking.indexOf(id) !== -1) usage.push(`${tag.name} (blocking)`);
  });
  data.list('trigger').forEach(trigger => {
    if (trigger.type !== 'triggerGroup') return;
    const ids = getListParamValues(trigger, 'triggerIds');
    if (ids.indexOf(id) !== -1) usage.push(`${trigger.name} (trigger group)`);
  });
  return usage;
}

/**
 * Column: tags that use the trigger (read-only).
 * @param {!Object=} options
 * @return {!Object}
 */
function triggerUsageColumn_(options) {
  return buildColumn_({
    key: 'usage',
    header: 'Used By',
    width: 320,
    kind: 'multiline',
    readOnly: true,
    note: 'Tags (and trigger groups) that use this trigger. Empty = unused.',
    get: (entity, data) => getTriggerUsage_(entity.triggerId, data).join('\n')
  }, options);
}
