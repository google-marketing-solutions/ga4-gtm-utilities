/**
 * Copyright 2022 Google LLC
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
 * Builds the menu.
 */
function onOpen() {
  const workspaceSubMenu = ui
  .createMenu('GTM Workspace Selection')
  .addItem('List GTM Accounts', 'writeGtmAccountsToSheet')
  .addItem('List GTM Containers', 'writeSelectedGtmContainersToSheet')
  .addItem('List GTM Workspaces', 'writeSelectedGtmWorkspacesToSheet')

  const modifyParamsAndUserProperties = ui
  .createMenu('Parameters and User Properties')
  .addItem('List Existing Parameters and User Properties', 'writeExistingParamsAndUserPropsToSheet')
  .addItem('List Existing GA4 Tags', 'writeGA4TagsToParamsAndUserPropertiesSheet')
  .addItem('Modify Parameters and User Properties', 'modifyParametersAndUserProperties');

  const eventTags = ui
  .createMenu('Event Tags')
  .addItem('List GA4 Event Tags', 'writeGA4EventTagstoSheet')
  .addItem('Modify GA4 Event Tags', 'modifyGA4EventTags');

  const listVariables = ui
  .createMenu('(Optional) GTM Variables')
  .addItem('Add GTM Variable Dropdown List Items', 'writeGTMVariablesToValidationSheet');

  const variableUsage = ui
  .createMenu('Variable Usage')
  .addItem('List', 'writeVariableUsageToSheet')
  .addItem('Delete', 'deleteUsageVariables');

  const sgtmSubMenu = ui
  .createMenu('Server-side Settings (sGTM)')
  .addItem('List Clients', 'writeSgtmClientsToSheet')
  .addItem('Modify Clients', 'modifySgtmClients')
  .addSeparator()
  .addItem('List Server Tags', 'writeSgtmTagsToSheet')
  .addItem('Modify Server Tags', 'modifySgtmTags');

  ui
  .createMenu('GA4 GTM Utilities')
  .addSubMenu(workspaceSubMenu)
  .addSubMenu(modifyParamsAndUserProperties)
  .addSubMenu(eventTags)
  .addSubMenu(sgtmSubMenu)
  .addSubMenu(variableUsage)
	.addSeparator()
  .addSubMenu(listVariables)
  .addToUi();
}