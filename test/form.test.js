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
  getLastColumn() {
    let maxCols = 0;
    this.grid.forEach(row => {
      if (row.length > maxCols) maxCols = row.length;
    });
    return maxCols || 1;
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
  let res = appScript.authenticateUser('admin', 'admin123');
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.username, 'admin');

  res = appScript.authenticateUser('admin', 'wrongpass');
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.message, 'Invalid Username or Password.');

  res = appScript.authenticateUser('', '');
  assert.strictEqual(res.success, false);
});

test('properCase capitalizes words correctly', () => {
  assert.strictEqual(appScript.properCase('john doe'), 'John Doe');
  assert.strictEqual(appScript.properCase('  saman PERERA '), 'Saman Perera');
  assert.strictEqual(appScript.properCase(''), '');
});

test('validateNIC validates Sri Lankan NIC numbers correctly', () => {
  assert.deepStrictEqual(appScript.validateNIC('200012345678'), { valid: true });
  assert.strictEqual(appScript.validateNIC('12345').valid, false);

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

test('saveTranscript and updateTranscript perform save and edit correctly', () => {
  activeSS = new MockSpreadsheet();
  appScript.setupSpreadsheet();

  // Save new
  let res = appScript.saveTranscript({ fullName: 'john doe', nicNumber: '912345678 V', batchNumber: '2019/A', nts: 'Colombo', date: '2023-01-01' });
  assert.strictEqual(res.success, true);

  // Get record for edit
  let editRes = appScript.getRecordForEdit('Transcript', '912345678 v');
  assert.strictEqual(editRes.found, true);
  assert.strictEqual(editRes.record.fullName, 'John Doe');

  // Update existing
  res = appScript.updateTranscript({ fullName: 'johnathan doe', nicNumber: '912345678 V', batchNumber: '2019/B', nts: 'Kandy', date: '2023-05-01' });
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.message, 'Transcript record updated successfully!');

  // Verify updated
  editRes = appScript.getRecordForEdit('Transcript', '912345678 V');
  assert.strictEqual(editRes.record.fullName, 'Johnathan Doe');
  assert.strictEqual(editRes.record.batchNumber, '2019/B');
});

test('saveExamResults and updateExamResults work properly', () => {
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

  let res = appScript.saveExamResults(data);
  assert.strictEqual(res.success, true);

  // Update
  data.paper1 = 'A+';
  res = appScript.updateExamResults(data);
  assert.strictEqual(res.success, true);

  const editRes = appScript.getRecordForEdit('Exam_Results', 'IND123');
  assert.strictEqual(editRes.found, true);
  assert.strictEqual(editRes.record.paper1, 'A+');
});

test('saveMediumLetter and updateMediumLetter work properly', () => {
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

  let res = appScript.saveMediumLetter(data);
  assert.strictEqual(res.success, true);

  data.nts = 'Colombo';
  res = appScript.updateMediumLetter(data);
  assert.strictEqual(res.success, true);

  const editRes = appScript.getRecordForEdit('Medium_Letter', '912345678 V');
  assert.strictEqual(editRes.record.nts, 'Colombo');
});

test('doGet, showHomePage and onOpen execute without error', () => {
  const htmlOutput = appScript.doGet();
  assert.ok(htmlOutput);

  appScript.showHomePage();
  assert.ok(activeUi.modalShown);

  appScript.onOpen();
  assert.ok(activeUi.menus.length > 0);
});
