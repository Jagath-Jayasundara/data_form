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
  assert.strictEqual(getEmployeeCategory('PHMA Grade I'), 'Senior');

  assert.strictEqual(getEmployeeCategory('S.K.S.'), 'Junior');
  assert.strictEqual(getEmployeeCategory('SKS'), 'Junior');
  assert.strictEqual(getEmployeeCategory('K.K.S.'), 'Junior');
  assert.strictEqual(getEmployeeCategory('KKS'), 'Junior');
});

test('validateAndNormalizeNic accepts valid old and new formats', () => {
  assert.strictEqual(validateAndNormalizeNic('123456789v'), '123456789V');
  assert.strictEqual(validateAndNormalizeNic('123456789V'), '123456789V');
  assert.strictEqual(validateAndNormalizeNic('200012345678'), '200012345678');

  // Invalid formats
  assert.throws(() => validateAndNormalizeNic('12345678V'), /valid NIC/);
  assert.throws(() => validateAndNormalizeNic('1234567890V'), /valid NIC/);
  assert.throws(() => validateAndNormalizeNic('123456789X'), /valid NIC/);
  assert.throws(() => validateAndNormalizeNic('12345678901'), /valid NIC/);
  assert.throws(() => validateAndNormalizeNic('1234567890123'), /valid NIC/);
});

test('normalizePhoneNumber strips non-numeric characters', () => {
  assert.strictEqual(normalizePhoneNumber('077-123-4567'), '0771234567');
  assert.strictEqual(normalizePhoneNumber('+94 77 123 4567'), '94771234567');
  assert.strictEqual(normalizePhoneNumber(''), '');
});

test('getFormOptions returns available capacity for locations', () => {
  const options = getFormOptions();
  assert.ok(Array.isArray(options.locations));
  assert.strictEqual(options.locations.length, Object.keys(LOCATION_CAPACITIES).length);

  const colombo = options.locations.find(l => l.name === 'PBCN-Colombo');
  assert.ok(colombo);
  assert.strictEqual(colombo.capacity, 9);
  assert.strictEqual(colombo.full, false);
});

test('submitForm successfully registers a Senior staff member', () => {
  const validData = {
    employeeName: 'Jane Doe',
    nic: '851234567v',
    designation: 'ICT Officer',
    workplace: 'Faculty of Science',
    telephone: '0771234567',
    preferredLocation: 'CON-Ampara'
  };

  const result = submitForm(validData);
  assert.strictEqual(result.success, true);
  assert.strictEqual(result.message, 'Your information has been submitted successfully.');

  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  const lastRow = sheet.rows[sheet.rows.length - 1];
  assert.strictEqual(lastRow[1], 'Jane Doe');
  assert.strictEqual(lastRow[2], '851234567V');
  assert.strictEqual(lastRow[3], 'ICT Officer');
  assert.strictEqual(lastRow[4], 'Senior');
  assert.strictEqual(lastRow[5], 'Faculty of Science');
  assert.strictEqual(lastRow[6], "'0771234567");
  assert.strictEqual(lastRow[7], 'CON-Ampara');
});

test('submitForm successfully registers a Junior staff member', () => {
  const validData = {
    employeeName: 'Perera',
    nic: '199512345678',
    designation: 'K.K.S.',
    workplace: 'Faculty of Arts',
    telephone: '0719876543',
    preferredLocation: 'CON-Kandy'
  };

  const result = submitForm(validData);
  assert.strictEqual(result.success, true);

  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  const lastRow = sheet.rows[sheet.rows.length - 1];
  assert.strictEqual(lastRow[1], 'Perera');
  assert.strictEqual(lastRow[2], '199512345678');
  assert.strictEqual(lastRow[3], 'K.K.S.');
  assert.strictEqual(lastRow[4], 'Junior');
  assert.strictEqual(lastRow[5], 'Faculty of Arts');
  assert.strictEqual(lastRow[6], "'0719876543");
  assert.strictEqual(lastRow[7], 'CON-Kandy');
});

test('submitForm prevents duplicate submission by same NIC', () => {
  const duplicateNicData = {
    employeeName: 'Jane Duplicate',
    nic: '851234567V',
    designation: 'Development Officer',
    workplace: 'Science',
    telephone: '0770000000',
    preferredLocation: 'PBCN-Colombo'
  };

  assert.throws(
    () => submitForm(duplicateNicData),
    /An entry with this NIC number has already been submitted/
  );
});

test('submitForm prevents duplicate submission by same telephone number', () => {
  const duplicatePhoneData = {
    employeeName: 'Jane Duplicate Phone',
    nic: '999999999V',
    designation: 'Development Officer',
    workplace: 'Science',
    telephone: '077-123-4567',
    preferredLocation: 'PBCN-Colombo'
  };

  assert.throws(
    () => submitForm(duplicatePhoneData),
    /An entry with this telephone number has already been submitted/
  );
});

test('submitForm rejects invalid NIC format', () => {
  const invalidNicData = {
    employeeName: 'Jane Doe',
    nic: '123456789',
    designation: 'Senior Lecturer',
    workplace: 'Faculty of Science',
    telephone: '0771234567',
    preferredLocation: 'CON-Ampara'
  };

  assert.throws(
    () => submitForm(invalidNicData),
    /Please enter a valid NIC number/
  );
});

test('submitForm validates phone number format', () => {
  const invalidPhoneData = {
    employeeName: 'Jane Doe',
    nic: '851234567V',
    designation: 'Senior Lecturer',
    workplace: 'Faculty of Science',
    telephone: 'abc',
    preferredLocation: 'CON-Ampara'
  };

  assert.throws(
    () => submitForm(invalidPhoneData),
    /Enter a valid telephone number/
  );
});

test('submitForm respects capacity limits', () => {
  const sheet = mockSpreadsheetInstance.getSheetByName('Responses');
  // Fill up CON-Ratnapura (capacity 2)
  sheet.appendRow([new Date(), 'A', '111111111V', 'S.K.S.', 'Junior', 'Uni', "'0771111111", 'CON-Ratnapura']);
  sheet.appendRow([new Date(), 'B', '222222222V', 'S.K.S.', 'Junior', 'Uni', "'0772222222", 'CON-Ratnapura']);

  const overflowData = {
    employeeName: 'Jane New',
    nic: '851234560V',
    designation: 'ICT Officer',
    workplace: 'Faculty of Science',
    telephone: '0778888888',
    preferredLocation: 'CON-Ratnapura'
  };

  assert.throws(
    () => submitForm(overflowData),
    /CON-Ratnapura has reached its maximum capacity of 2/
  );
});
