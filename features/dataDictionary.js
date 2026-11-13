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
 * Formats GTM triggers.
 */
function formatTriggers() {
  const triggers = listGTMResources('triggers', getSelectedWorkspacePath());
  if (triggers.length > 0) {
    triggers.forEach(trigger => {
      console.log(trigger);
    });
  }
}

/**
 * Formats GTM tags.
 */
function formatTags() {
  const workspace = getSelectedWorkspacePath();
  const tags = listGTMResources('tags', workspace);
  const triggers = listGTMResources('triggers', workspace);
  const variables = listGTMResources('variables', workspace);
  let finalRows = [];
  if (tags.length > 0) {
    finalRows = [...tags.reduce((tagArray, tag) => {
      let variableNames = [];
      if (tag.parameter != undefined) {
        variableNames = [...tag.parameter.reduce((paramArray, param) => {
          if (/{{/.test(param.value)) {
            const cleanVarName = param.value.split('{{')[1].split('}}')[0];
            const filteredVariable = variables.filter(
              variable => variable.name == cleanVarName);
            if (filteredVariable.length > 0) {
              paramArray.push(cleanVarName);
            }
          }
          return paramArray;
        }, [])];
      }
      let firingTriggers = '';
      if (tag.firingTriggerId != undefined) {
        firingTriggers = convertTriggerIdsToNames(
          tag.firingTriggerId, triggers).join(', ');
      }
      let blockingTriggers = '';
      if (tag.blockingTriggerId != undefined) {
        blockingTriggers = convertTriggerIdsToNames(
          tag.blockingTriggerId, triggers).join(', ');
      }
      tagArray.push([
        buildHyperlinkFormula(tag.tagManagerUrl, tag.name),
        tag.tagId,
        tag.type,
        firingTriggers,
        blockingTriggers,
        variableNames,
        tag.notes
      ]);
      return tagArray;
    }, [])];
  }
  writeToSheet(finalRows, 'tagDataDictionary', 'tag info');
}