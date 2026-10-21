const test = require('node:test');
const assert = require('node:assert');

// Mock Google Apps Script environment
class MockRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  merge() { return this; }
  setValue(val) {
    this.sheet.grid[this.row - 1] = this.sheet.grid[this.row - 1] || [];
    this.sheet.grid[this.row - 1][this.col - 1] = val;
    return this;
  }
  setValues(vals) {
    for (let r = 0; r < vals.length; r++) {
      const targetRow = this.row - 1 + r;
      this.sheet.grid[targetRow] = this.sheet.grid[targetRow] || [];
      for (let c = 0; c < vals[r].length; c++) {
        this.sheet.grid[targetRow][this.col - 1 + c] = vals[r][c];
      }
    }
    return this;
  }
  getValues() {
    const result = [];
    for (let r = 0; r < this.numRows; r++) {
      const targetRow = this.row - 1 + r;
      const rowData = [];
      for (let c = 0; c < this.numCols; c++) {
        const val = (this.sheet.grid[targetRow] && this.sheet.grid[targetRow][this.col - 1 + c] !== undefined)
          ? this.sheet.grid[targetRow][this.col - 1 + c]
          : '';
        rowData.push(val);
      }
      result.push(rowData);
    }
    return result;
  }
  setFontWeight() { return this; }
  setFontSize() { return this; }
  setHorizontalAlignment() { return this; }
  setBackground() { return this; }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.grid = [];
  }
  clear() {
    this.grid = [];
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      return new MockRange(this, 1, 1, 1, 1);
    }
    return new MockRange(this, a, b, c, d);
  }
  getLastRow() {
    return this.grid.length;
  }
  setFrozenRows() {}
  appendRow(rowArray) {
    this.grid.push([...rowArray]);
  }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = [];
  }
  getSheetByName(name) {
    return this.sheets.find(s => s.name === name) || null;
  }
  insertSheet(name, index) {
    const sheet = new MockSheet(name);
    if (typeof index === 'number') {
      this.sheets.splice(index, 0, sheet);
    } else {
      this.sheets.push(sheet);
    }
    return sheet;
  }
}

class MockUi {
  constructor() {
    this.menus = [];
    this.modalShown = null;
  }
  createMenu(title) {
    const menu = { title, items: [] };
    this.menus.push(menu);
    return {
      addItem: (itemTitle, fnName) => {
        menu.items.push({ itemTitle, fnName });
        return this.createMenu(title);
      },
      addSeparator: () => this.createMenu(title),
      addToUi: () => {}
    };
  }
  showModalDialog(html, title) {
    this.modalShown = { html, title };
  }
}

class MockHtmlOutput {
  constructor(filename) {
    this.filename = filename;
    this.title = '';
    this.width = 0;
    this.height = 0;
  }
  setTitle(title) { this.title = title; return this; }
  setXFrameOptionsMode() { return this; }
  setWidth(w) { this.width = w; return this; }
  setHeight(h) { this.height = h; return this; }
}

let activeSS = new MockSpreadsheet();
let activeUi = new MockUi();

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => activeSS,
  getUi: () => activeUi
};

global.HtmlService = {
  createHtmlOutputFromFile: (filename) => new MockHtmlOutput(filename),
  XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
};

const appScript = require('../Code.gs');

test('authenticateUser validates credentials correctly', () => {
  // Valid
  let res = appScript.authenticateUser('admin', 'admin123');
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.username, 'admin');

  // Invalid
  res = appScript.authenticateUser('admin', 'wrongpass');
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.message, 'Invalid Username or Password.');

  // Empty
  res = appScript.authenticateUser('', '');
  assert.strictEqual(res.success, false);
});

test('properCase capitalizes words correctly', () => {
  assert.strictEqual(appScript.properCase('john doe'), 'John Doe');
  assert.strictEqual(appScript.properCase('  saman PERERA '), 'Saman Perera');
  assert.strictEqual(appScript.properCase(''), '');
});

test('validateNIC validates Sri Lankan NIC numbers correctly', () => {
  // New NIC
  assert.deepStrictEqual(appScript.validateNIC('200012345678'), { valid: true });
  assert.strictEqual(appScript.validateNIC('12345').valid, false);

  // Old NIC
  assert.deepStrictEqual(appScript.validateNIC('912345678 V'), { valid: true });
  assert.deepStrictEqual(appScript.validateNIC('912345678 v'), { valid: true });
  assert.strictEqual(appScript.validateNIC('912345678V').valid, false);
  assert.strictEqual(appScript.validateNIC('ABC345678 V').valid, false);
  assert.strictEqual(appScript.validateNIC('912345678 X').valid, false);
});

test('setupSpreadsheet creates Home and data sheets with headers', () => {
  activeSS = new MockSpreadsheet();
  const res = appScript.setupSpreadsheet();
  assert.strictEqual(res, 'Spreadsheet setup completed successfully!');

  assert.ok(activeSS.getSheetByName('Home'));
  assert.ok(activeSS.getSheetByName('Transcript'));
  assert.ok(activeSS.getSheetByName('Exam_Results'));
  assert.ok(activeSS.getSheetByName('Medium_Letter'));
});

test('saveTranscript validates inputs and prevents duplicate NICs', () => {
  activeSS = new MockSpreadsheet();
  appScript.setupSpreadsheet();

  // Missing field
  let res = appScript.saveTranscript({ fullName: 'John', nicNumber: '', batchNumber: '2019/A', nts: 'Colombo', date: '2023-01-01' });
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.message, 'Please fill all text fields');

  // Invalid NIC
  res = appScript.saveTranscript({ fullName: 'John Doe', nicNumber: '123', batchNumber: '2019/A', nts: 'Colombo', date: '2023-01-01' });
  assert.strictEqual(res.success, false);

  // Successful save
  res = appScript.saveTranscript({ fullName: 'john doe', nicNumber: '912345678 V', batchNumber: '2019/A', nts: 'Colombo', date: '2023-01-01' });
  assert.strictEqual(res.success, true);

  // Duplicate NIC save
  res = appScript.saveTranscript({ fullName: 'Jane Doe', nicNumber: '912345678 V', batchNumber: '2020/B', nts: 'Kandy', date: '2023-01-02' });
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.message, 'This NIC number already exists. Please enter a unique NIC.');
});

test('lookupStudentByNic finds registered students in Transcript sheet', () => {
  activeSS = new MockSpreadsheet();
  appScript.setupSpreadsheet();
  appScript.saveTranscript({ fullName: 'saman kumara', nicNumber: '912345678 V', batchNumber: '2018/A', nts: 'Galle', date: '2023-01-01' });

  // Found
  let result = appScript.lookupStudentByNic('912345678 v');
  assert.strictEqual(result.found, true);
  assert.strictEqual(result.fullName, 'Saman Kumara');
  assert.strictEqual(result.batchNumber, '2018/A');
  assert.strictEqual(result.nts, 'Galle');

  // Not found
  result = appScript.lookupStudentByNic('000000000 V');
  assert.strictEqual(result.found, false);
  assert.strictEqual(result.message, 'Student not found');
});

test('saveExamResults saves valid exam records', () => {
  activeSS = new MockSpreadsheet();
  appScript.setupSpreadsheet();

  const data = {
    title: 'Mr.',
    fullName: 'Saman Kumara',
    indexNumber: 'IND123',
    nic: '912345678 V',
    batchNumber: '2019/A',
    nts: 'Galle',
    hisHer: 'his',
    fromDate: '2019-01-01',
    toDate: '2022-01-01',
    paper1: 'A',
    paper2: 'B',
    paper3: 'A',
    paper4: 'B',
    paper5: 'A',
    practical: 'Pass',
    date: '2023-01-01'
  };

  const res = appScript.saveExamResults(data);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.message, 'Data saved successfully!');

  const examSheet = activeSS.getSheetByName('Exam_Results');
  assert.strictEqual(examSheet.getLastRow(), 2);
});

test('saveMediumLetter saves valid medium letter records', () => {
  activeSS = new MockSpreadsheet();
  appScript.setupSpreadsheet();

  const data = {
    nic: '912345678 V',
    title: 'Mr.',
    fullName: 'Saman Kumara',
    batchNo: '2019/A',
    nts: 'Galle',
    fromDate: '2019-01-01',
    toDate: '2022-01-01',
    date: '2023-01-01'
  };

  const res = appScript.saveMediumLetter(data);
  assert.strictEqual(res.success, true);

  const sheet = activeSS.getSheetByName('Medium_Letter');
  assert.strictEqual(sheet.getLastRow(), 2);
});

test('doGet, showHomePage and onOpen execute without error', () => {
  const htmlOutput = appScript.doGet();
  assert.ok(htmlOutput);

  appScript.showHomePage();
  assert.ok(activeUi.modalShown);

  appScript.onOpen();
  assert.ok(activeUi.menus.length > 0);
});
