const test = require('node:test');
const assert = require('node:assert');

class MockRange {
  constructor(values = []) {
    this.values = values;
  }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
  getDisplayValues() { return this.values; }
}

class MockSheet {
  constructor(name = 'Responses') {
    this.name = name;
    this.rows = [];
  }
  setName(name) {
    this.name = name;
  }
  getLastRow() {
    return this.rows.length;
  }
  appendRow(row) {
    this.rows.push(row);
  }
  setFrozenRows() {}
  getRange(row, col, numRows = 1, numCols = 1) {
    if (typeof row === 'number') {
      const slice = this.rows.slice(row - 1, row - 1 + numRows);
      const values = slice.map(r => r.slice(col - 1, col - 1 + numCols));
      return new MockRange(values);
    }
    return new MockRange();
  }
  autoResizeColumns() {}
}

class MockSpreadsheet {
  constructor(id = 'mock-sheet-id') {
    this.id = id;
    this.sheets = [new MockSheet('Responses')];
  }
  getId() {
    return this.id;
  }
  getUrl() {
    return `https://docs.google.com/spreadsheets/d/${this.id}/edit`;
  }
  getSheetByName(name) {
    return this.sheets.find(s => s.name === name) || null;
  }
  getSheets() {
    return this.sheets;
  }
}

const mockPropertiesStore = {
  EXAM_STAFF_SPREADSHEET_ID: 'mock-sheet-id'
};

const mockSpreadsheetInstance = new MockSpreadsheet('mock-sheet-id');

global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (key) => mockPropertiesStore[key] || null,
    setProperty: (key, val) => { mockPropertiesStore[key] = val; }
  })
};

global.SpreadsheetApp = {
  openById: (id) => mockSpreadsheetInstance,
  create: (name) => mockSpreadsheetInstance
};

global.LockService = {
  getScriptLock: () => ({
    waitLock: () => {},
    releaseLock: () => {}
  })
};

global.HtmlService = {
  createHtmlOutputFromFile: (file) => ({
    setTitle: () => ({ addMetaTag: () => ({}) })
  })
};

global.Logger = {
  log: () => {}
};

const {
  LOCATION_CAPACITIES,
  JUNIOR_LOCATION_CAPACITIES,
  setup,
  doGet,
  getFormOptions,
  submitForm,
  getEmployeeCategory,
  normalizePhoneNumber,
  validateAndNormalizeNic,
  cleanField
} = require('../Code.gs');

test('setup initializes spreadsheet headers correctly with Category and NIC', () => {
  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  sheet.rows = [];
  const url = setup();
  assert.ok(url.includes('mock-sheet-id'));
  assert.strictEqual(sheet.rows.length, 1);
  assert.deepStrictEqual(sheet.rows[0], [
    'Submission Date',
    'Employee Name',
    'NIC Number',
    'Designation',
    'Category',
    'Current Workplace',
    'Telephone Number',
    'Preferred Examination Duty Location'
  ]);
});

test('getEmployeeCategory classifies Senior and Junior roles correctly', () => {
  assert.strictEqual(getEmployeeCategory('ICT Officer'), 'Senior');
  assert.strictEqual(getEmployeeCategory('Development Officer'), 'Senior');
  assert.strictEqual(getEmployeeCategory('Public Health Management Assistant - Grade I'), 'Senior');
  assert.strictEqual(getEmployeeCategory('Public Health Management Assistant - Grade II'), 'Senior');
  assert.strictEqual(getEmployeeCategory('Public Health Management Assistant - Grade III'), 'Senior');
  assert.strictEqual(getEmployeeCategory('S.K.S.'), 'Junior');
  assert.strictEqual(getEmployeeCategory('K.K.S.'), 'Junior');
});

test('junior registration does not decrement Senior vacancies', () => {
  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  sheet.rows = [];
  setup();

  // Initial Senior options
  const initialSeniorOptions = getFormOptions('Senior');
  const initialSeniorColombo = initialSeniorOptions.locations.find(l => l.name === 'PBCN-Colombo');
  assert.strictEqual(initialSeniorColombo.remaining, 9);

  // Submit a Junior registration
  const juniorSubmission = {
    employeeName: 'Junior A',
    nic: '111111111V',
    designation: 'S.K.S.',
    workplace: 'Uni',
    telephone: '0771111111',
    preferredLocation: 'PBCN-Colombo'
  };
  submitForm(juniorSubmission);

  // Check Junior options: should be 1 remaining out of 2
  const juniorOptions = getFormOptions('Junior');
  const juniorColombo = juniorOptions.locations.find(l => l.name === 'PBCN-Colombo');
  assert.strictEqual(juniorColombo.remaining, 1);

  // Check Senior options: should still be 9 remaining out of 9
  const seniorOptionsAfter = getFormOptions('Senior');
  const seniorColomboAfter = seniorOptionsAfter.locations.find(l => l.name === 'PBCN-Colombo');
  assert.strictEqual(seniorColomboAfter.remaining, 9);
});

test('submitForm respects Junior quota of 2 vacancies', () => {
  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  sheet.rows = [];
  setup();

  const junior1 = {
    employeeName: 'Junior A',
    nic: '111111111V',
    designation: 'S.K.S.',
    workplace: 'Uni',
    telephone: '0771111111',
    preferredLocation: 'PBCN-Colombo'
  };
  const junior2 = {
    employeeName: 'Junior B',
    nic: '222222222V',
    designation: 'K.K.S.',
    workplace: 'Uni',
    telephone: '0772222222',
    preferredLocation: 'PBCN-Colombo'
  };
  const junior3 = {
    employeeName: 'Junior C',
    nic: '333333333V',
    designation: 'S.K.S.',
    workplace: 'Uni',
    telephone: '0773333333',
    preferredLocation: 'PBCN-Colombo'
  };

  assert.strictEqual(submitForm(junior1).success, true);
  assert.strictEqual(submitForm(junior2).success, true);

  assert.throws(
    () => submitForm(junior3),
    /PBCN-Colombo has reached its maximum capacity of 2 staff members for the Junior quota/
  );
});

test('submitForm prevents duplicate submission by same NIC', () => {
  const duplicateNicData = {
    employeeName: 'Junior Duplicate',
    nic: '111111111V',
    designation: 'S.K.S.',
    workplace: 'Uni',
    telephone: '0779999999',
    preferredLocation: 'CON-Ampara'
  };

  assert.throws(
    () => submitForm(duplicateNicData),
    /An entry with this NIC number has already been submitted/
  );
});
