/**
 * Copyright 2023 Google LLC
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
 * Writes variable usage data to the sheet.
 */
function writeVariableUsageToSheet() {
  clearRangeContent('variableUsage', sheetsMeta.variableUsage.ranges[0].name);
  const variableUsageData = getAllVariableUsage();
  writeToSheet(
    variableUsageData, 'variableUsage',
    sheetsMeta.variableUsage.ranges[0].name);
  const usageSheet = ss.getSheetByName(sheetsMeta.variableUsage.sheetName);
  if (variableUsageData.length > 0) {
    usageSheet.setRowHeightsForced(2, variableUsageData.length, 21);
  }
}

/**
 * Deletes variables marked for deletion.
 */
function deleteUsageVariables() {
  const usageSheetData = getDataFromSheet(
    'variableUsage', sheetsMeta.variableUsage.ranges[0].name);
  if (usageSheetData) {
    usageSheetData.forEach(variable => {
      if (variable[variable.length - 1]) {
        const path = variable[1];
        const url = `https://tagmanager.google.com/#/container/${path}`;
        try {
          removeGTMResource('variables', path);
          logChange(
            variable[0], 'variable', path.split('/')[7], 'Deleted', url);
        } catch(e) {
          logChange(
            variable[0], 'variable', path.split('/')[7], e.message, url);
        }
      }
    });
  }
}

/**
 * Gets all variables and searches for their usage in tags, triggers, other variables, and clients.
 * @return {!Array<!Array>} Array of variable usage details.
 */
function getAllVariableUsage() {
  const workspace = getSelectedWorkspacePath();
  const variables = listGTMResources('variables', workspace);
  const tags = listGTMResources('tags', workspace);
  const triggers = listGTMResources('triggers', workspace);
  
  const isServer = isServerContainer(workspace);
  const clients = isServer ? listGTMResources('clients', workspace) : null;

  const finalData = [];
  if (variables) {
    variables.forEach(variable => {
      const tagInfo = getVariableUsageForTags(tags, variable);
      const triggerInfo = getVariableUsageForTriggers(triggers, variable);
      const variableInfo = getVariableUsageForVariables(variables, variable);
      const clientInfo = isServer ? getVariableUsageForClients(clients, variable) : { clientUsageCount: 0, clientNames: '' };
      
      finalData.push([
        `=HYPERLINK("https://tagmanager.google.com/#/container/${variable.path}","${variable.name}")`,
        variable.path,
        tagInfo.tagUsageCount, 
        tagInfo.tagNames,
        triggerInfo.triggerUsageCount,
        triggerInfo.triggerNames,
        variableInfo.variableUsageCount,
        variableInfo.variableNames,
        clientInfo.clientUsageCount,
        clientInfo.clientNames
      ]);
    });
  }
  return finalData;
}

/**
 * Checks if a variable is inside parameters list.
 * @param {string} variableName Name of variable.
 * @param {!Array} parameters Parameters list.
 * @return {boolean} True if found.
 */
function checkIfVariableIsInParameters(variableName, parameters) {
  const variableNameWithBrackets = `{{${variableName}}}`;
  if (parameters) {
    for (let i = 0; i < parameters.length; i++) {
      const param = parameters[i];
      if (param.value !== undefined && param.value !== null) {
        const paramValueStr = String(param.value);
        if (param.key != 'html' && param.key != 'javascript') {
          if (paramValueStr.indexOf(variableNameWithBrackets) !== -1) {
            return true;
          }
        } else {
          const regexVariableName = new RegExp(variableNameWithBrackets);
          if (regexVariableName.test(paramValueStr)) {
            return true;
          }
        }
      }
      if (param.list && checkIfVariableIsInParameters(variableName, param.list)) {
        return true;
      } 
      if (param.map && checkIfVariableIsInParameters(variableName, param.map)) {
        return true;
      }
      if (param.parameter && checkIfVariableIsInParameters(variableName, param.parameter)) {
        return true;
      }
    } 
  }
  return false;
}

/**
 * Searches variable usage inside tags.
 * @param {!Array} tags Tags list.
 * @param {!Object} variable Variable object.
 * @return {!Object} Usage details.
 */
function getVariableUsageForTags(tags, variable) {
  const tagInfo = {
    tagNames: '',
    tagUsageCount: 0
  };
  if (tags) {
    tags.forEach(tag => {
      const isInTag = 
        checkIfVariableIsInParameters(variable.name, tag.parameter) ||
        checkIfVariableIsInParameters(variable.name, tag.priority) ||
        checkIfVariableIsInParameters(
          variable.name, tag.monitoringMetadata) ||
        checkIfVariableIsInParameters(
          variable.name, tag.consentSettings.consentType);
      if (isInTag) {
        tagInfo.tagUsageCount++;
        if (tagInfo.tagNames) {
          tagInfo.tagNames += `,\n${tag.name}`;
        } else {
          tagInfo.tagNames += tag.name;
        }
      }
    });
  }
  return tagInfo;
}

/**
 * Searches variable usage inside triggers.
 * @param {!Array} triggers Triggers list.
 * @param {!Object} variable Variable object.
 * @return {!Object} Usage details.
 */
function getVariableUsageForTriggers(triggers, variable) {
  const triggerInfo = {
    triggerNames: '',
    triggerUsageCount: 0
  };
  if (triggers) {
    triggers.forEach(trigger => {
      let isInTrigger = 
        checkIfVariableIsInParameters(variable.name, trigger.parameter);
      const fields = [
        'customEvent',
        'filter',
        'autoEventFilter',
        'waitForTags',
        'checkValidation',
        'waitForTagsTimeout',
        'uniqueTriggerId',
        'eventName',
        'interval',
        'limit',
        'selector',
        'intervalSeconds',
        'maxTimerLengthSeconds',
        'verticalScrollPercentageList',
        'horizontalScrollPercentageList',
        'visibilitySelector',
        'visiblePercentageMin',
        'visiblePercentageMax',
        'continuousTimeMinMilliseconds',
        'totalTimeMinMilliseconds'
      ];
      fields.forEach(field => {
        if (trigger[field]) {
          if (Array.isArray(trigger[field])) {
            if (checkIfVariableIsInParameters(
              variable.name, trigger[field])) {
              isInTrigger = true;
            }
          } else {
            if (checkIfVariableIsInParameters(
              variable.name, [trigger[field]])) {
              isInTrigger = true;
            }
          }
        }
      });
      if (isInTrigger) {
        triggerInfo.triggerUsageCount++;
        if (triggerInfo.triggerNames) {
          triggerInfo.triggerNames += `,\n${trigger.name}`;
        } else {
          triggerInfo.triggerNames += trigger.name;
        }
      }
    });
  }
  return triggerInfo;
}

/**
 * Searches variable usage inside other variables.
 * @param {!Array} variables Variables list.
 * @param {!Object} variable Target variable object.
 * @return {!Object} Usage details.
 */
function getVariableUsageForVariables(variables, variable) {
  const variableInfo = {
    variableNames: '',
    variableUsageCount: 0
  };
  variables.forEach(v => {
    let isInVariable = 
        checkIfVariableIsInParameters(variable.name, v.parameter);
    const fields = [
      'convertNullToValue',
      'convertUndefinedToValue',
      'convertTrueToValue',
      'convertFalseToValue'
    ];
    if (v.formatValue) {
      fields.forEach(field => {
        if (v.formatValue[field]) {
          if (checkIfVariableIsInParameters(
            variable.name, [v.formatValue[field]])) {
            isInVariable = true;
          }
        }
      });
    }
    if (isInVariable) {
      variableInfo.variableUsageCount++;
      if (variableInfo.variableNames) {
        variableInfo.variableNames += `,\n${v.name}`;
      } else {
        variableInfo.variableNames += v.name;
      }
    }
  });
  return variableInfo;
}

/**
 * Searches variable usage inside server-side clients.
 * @param {!Array} clients Clients list.
 * @param {!Object} variable Target variable object.
 * @return {!Object} Usage details.
 */
function getVariableUsageForClients(clients, variable) {
  const clientInfo = {
    clientNames: '',
    clientUsageCount: 0
  };
  if (clients) {
    clients.forEach(client => {
      const isInClient = checkIfVariableIsInParameters(variable.name, client.parameter);
      if (isInClient) {
        clientInfo.clientUsageCount++;
        if (clientInfo.clientNames) {
          clientInfo.clientNames += `,\n${client.name}`;
        } else {
          clientInfo.clientNames += client.name;
        }
      }
    });
  }
  return clientInfo;
}
