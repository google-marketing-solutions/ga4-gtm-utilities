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

const sheetsMeta = {
  gtmWorkspace: {
    sheetName: 'GTM Workspace',
    ranges: [{
      name: 'accounts list',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 3
      },
      write : {
        row: 2,
        column: 1,
        numRow: 1,
        numColumns: 2
      }
    }, {
      name: 'containers list',
      read: {
        row: 2,
        column: 4,
        numRows: 1,
        numColumns: 3
      },
      write : {
        row: 2,
        column: 4,
        numRow: 1,
        numColumns: 2
      }
    }, {
      name: 'workspaces list',
      read: {
        row: 2,
        column: 7,
        numRows: 1,
        numColumns: 3
      },
      write : {
        row: 2,
        column: 7,
        numRow: 1,
        numColumns: 2
      }
    }]
  },
  modifyParamsAndUserProps: {
    sheetName: 'Modify Parameter or User Property Settings',
    ranges: [{
      name: 'settings',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 8
      },
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 6
      }
    }]
  },
  validation: {
    sheetName: 'Validation Settings',
    ranges: [{
      name: 'gtm variables',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 1
      },
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 1
      }
    }, {
      name: 'tag names',
      read: {
        row: 2,
        column: 4,
        numRows: 1,
        numColumns: 1
      },
      write: {
        row: 2,
        column: 4,
        numRows: 1,
        numColumns: 1
      }
    }, {
      name: 'event setting variable names',
      read: {
        row: 2,
        column: 3,
        numRows: 1,
        numColumns: 1
      },
      write: {
        row: 2,
        column: 3,
        numRows: 1,
        numColumns: 1
      }
    }]
  },
  eventTags: {
    sheetName: 'Event Tag Settings',
    ranges: [{
      name: 'tags',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 30
      }, 
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 26
      }
    }]
  },
  tagDataDictionary: {
    sheetName: 'Tag Data Dictionary',
    ranges: [{
      name: 'tag info',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 7
      }, 
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 7
      }
    }]
  },
  settings: {
    sheetName: 'Settings',
    ranges: [{
      name: 'dismissed update',
      read: {
        row: 2,
        column: 2,
        numRows: 1,
        numColumns: 1
      },
      write: {
        row: 2,
        column: 2,
        numRows: 1,
        numColumns: 1
      }
    }, {
      name: 'release',
      read: {
        row: 1,
        column: 2,
        numRows: 1,
        numColumns: 1
      },
      write: {
        row: 1,
        column: 2,
        numRows: 1,
        numColumns: 1
      }
    }]
  },
  variableUsage: {
    sheetName: 'Variable Usage',
    ranges: [{
      name: 'all usage',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 10
      },
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 8
      }
    }]
  },
  sgtmClients: {
    sheetName: '(sGTM) Client Settings',
    ranges: [{
      name: 'clients',
      read: {
        row: 3,
        column: 1,
        numRows: 1,
        numColumns: 26
      },
      write: {
        row: 3,
        column: 1,
        numRows: 1,
        numColumns: 22
      }
    }]
  },
  sgtmTags: {
    sheetName: '(sGTM) Tag Settings',
    ranges: [{
      name: 'tags',
      read: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 18
      },
      write: {
        row: 2,
        column: 1,
        numRows: 1,
        numColumns: 14
      }
    }]
  }
};