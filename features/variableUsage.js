/**
 * 
 */
function writeVariableUsageToSheet() {
  clearRangeContent('variableUsage', sheetsMeta.variableUsage.ranges[0].name);
  const variableUsageData = getAllVariableUsage();
  writeToSheet(
    variableUsageData,'variableUsage',
    sheetsMeta.variableUsage.ranges[0].name);
  const usageSheet = ss.getSheetByName(sheetsMeta.variableUsage.sheetName);
  usageSheet.setRowHeightsForced(2, variableUsageData.length, 21);  
}

/**
 * 
 */
function deleteUsageVariables() {
  const usageSheetData = getDataFromSheet(
    'variableUsage', sheetsMeta.variableUsage.ranges[0].name);
  if (usageSheetData) {
    usageSheetData.forEach(variable => {
      if (variable[variable.length - 1]) {
        const url = 
          `https://tagmanager.google.com/#/container/${variable.path}`;
        const path = variable[1];
        try {
          removeGTMResource('variables', path);
          logChange(
            variable[0], 'variable', variable.split('/')[7], 'Deleted', url);
        } catch(e) {
          logChange(
            variable[0], 'variable', variable.split('/')[7], e.message, url);
        }
      }
    });
  }
}

/**
 * 
 */
function getAllVariableUsage() {
  const workspace = getSelectedWorkspacePath();
  const variables = listGTMResources('variables', workspace);
  const tags = listGTMResources('tags', workspace);
  const triggers = listGTMResources('triggers', workspace);
  const finalData = [];
  if (variables) {
    variables.forEach(variable => {
      const tagInfo = getVariableUsageForTags(tags, variable);
      const triggerInfo = getVariableUsageForTriggers(triggers, variable);
      const variableInfo = getVariableUsageForVariables(variables, variable);
      finalData.push([
        `=HYPERLINK("https://tagmanager.google.com/#/container/${variable.path}","${variable.name}")`,
        variable.path,
        tagInfo.tagUsageCount, 
        tagInfo.tagNames,
        triggerInfo.triggerUsageCount,
        triggerInfo.triggerNames,
        variableInfo.variableUsageCount,
        variableInfo.variableNames
      ]);
    });
  }
  return finalData;
}

/**
 * 
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
 * 
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
 * 
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
 * 
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
