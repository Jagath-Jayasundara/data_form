/**
 * Authenticates user credentials for accessing the Student Management System.
 * Default credentials: username "admin", password "admin123".
 * @param {string} username - User username.
 * @param {string} password - User password.
 * @return {{success: boolean, username?: string, message?: string}}
 */
function authenticateUser(username, password) {
  if (!username || !password) {
    return { success: false, message: 'Please enter both Username and Password.' };
  }

  const u = String(username).trim();
  const p = String(password).trim();

  if (u === 'admin' && p === 'admin123') {
    return { success: true, username: u };
  } else {
    return { success: false, message: 'Invalid Username or Password.' };
  }
}

/**
 * Capitalizes the first letter of each word in a string.
 * @param {string} str - The string to convert.
 * @return {string} The string in Proper Case.
 */
function properCase(str) {
  if (!str) return '';
  const arr = str.trim().toLowerCase().split(/\s+/);
  for (let i = 0; i < arr.length; i++) {
    if (arr[i].length > 0) {
      arr[i] = arr[i].charAt(0).toUpperCase() + arr[i].slice(1);
    }
  }
  return arr.join(' ');
}

/**
 * Validates Sri Lankan NIC (National Identity Card) format.
 * - New NIC: Exactly 12 digits.
 * - Old NIC: 11 characters, first 9 digits, followed by a space, ending with 'V' or 'v' (e.g., "912345678 V").
 * @param {string} nic - The NIC string.
 * @return {{valid: boolean, message?: string}} Validation result.
 */
function validateNIC(nic) {
  if (!nic) {
    return { valid: false, message: 'NIC is required.' };
  }
  const trimmed = String(nic).trim();

  if (/^\d+$/.test(trimmed)) {
    if (trimmed.length !== 12) {
      return { valid: false, message: 'New NIC number must contain exactly 12 digits.' };
    }
    return { valid: true };
  } else {
    if (trimmed.length !== 11) {
      return { valid: false, message: 'Old NIC must be in the format: 912345678 V' };
    }
    const first9 = trimmed.substring(0, 9);
    if (!/^\d{9}$/.test(first9)) {
      return { valid: false, message: 'The first 9 characters must be digits.' };
    }
    if (trimmed.charAt(9) !== ' ') {
      return { valid: false, message: 'Please enter a space before the letter V.' };
    }
    if (trimmed.charAt(10).toUpperCase() !== 'V') {
      return { valid: false, message: 'Old NIC must end with the letter V.' };
    }
    return { valid: true };
  }
}

/**
 * Returns list of Nursing Training Schools (NTS).
 * @return {string[]} Array of NTS options.
 */
function getNTSList() {
  return [
    'Ampara',
    'Anuradhapura',
    'Anuradhapura (Military)',
    'Badulla',
    'Batticaloa',
    'Colombo',
    'Galle',
    'Hambanthota',
    'Jaffna',
    'Kandy',
    'Kandana',
    'Kalutara',
    'Kurunegala',
    'Matara',
    'Mulleriyawa',
    'Rathnapura',
    'Vavuniya',
    'Sri Jayawardenepura'
  ];
}

/**
 * Returns list of Title options.
 * @return {string[]} Array of Title options.
 */
function getTitleList() {
  return ['Mr.', 'Ms.', 'Miss.'];
}

/**
 * Returns list of His/Her options.
 * @return {string[]} Array of His/Her options.
 */
function getHisHerList() {
  return ['his', 'her'];
}

/**
 * Helper to get or create a sheet in the active spreadsheet.
 * @param {string} sheetName - The sheet name.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Sets up spreadsheet with all required sheets and headers.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Home Sheet
  let homeSheet = ss.getSheetByName('Home');
  if (!homeSheet) {
    homeSheet = ss.insertSheet('Home', 0);
  }
  homeSheet.clear();
  homeSheet.getRange('A1:D1').merge().setValue('STUDENT MANAGEMENT SYSTEM - HOME PAGE ACCESS TAB');
  homeSheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  homeSheet.getRange('A3').setValue('Available Forms & Modules:');
  homeSheet.getRange('A3').setFontWeight('bold');

  const modules = [
    ['Module Name', 'Target Sheet', 'Description'],
    ['1. Transcript / Student Entry', 'Transcript', 'Register new student transcript details & validate NIC'],
    ['2. Exam Results Entry', 'Exam_Results', 'Record exam paper results for registered students'],
    ['3. Medium Letter Entry', 'Medium_Letter', 'Issue medium of instruction letters for students']
  ];
  homeSheet.getRange(4, 1, modules.length, 3).setValues(modules);
  homeSheet.getRange('A4:C4').setFontWeight('bold').setBackground('#E0E0E0');

  // Setup Transcript Sheet
  const transcriptSheet = getOrCreateSheet('Transcript');
  if (transcriptSheet.getLastRow() === 0) {
    const headers = [['Full Name', 'NIC Number', 'Batch Number', 'NTS', 'Date']];
    transcriptSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#E0E0E0');
    transcriptSheet.setFrozenRows(1);
  }

  // Setup Exam Results Sheet
  const examSheet = getOrCreateSheet('Exam_Results');
  if (examSheet.getLastRow() === 0) {
    const headers = [[
      'Title', 'Full Name', 'Index Number', 'NIC', 'Batch Number',
      'NTS', 'His/Her', 'From Date', 'To Date',
      'Paper I', 'Paper II', 'Paper III', 'Paper IV', 'Paper V', 'Practical', 'Date'
    ]];
    examSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#E0E0E0');
    examSheet.setFrozenRows(1);
  }

  // Setup Medium Letter Sheet
  const mediumSheet = getOrCreateSheet('Medium_Letter');
  if (mediumSheet.getLastRow() === 0) {
    const headers = [['Title', 'Full Name', 'NIC', 'Batch Number', 'NTS', 'From Date', 'To Date', 'Date']];
    mediumSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold').setBackground('#E0E0E0');
    mediumSheet.setFrozenRows(1);
  }

  return 'Spreadsheet setup completed successfully!';
}

/**
 * Searches for a student by NIC in the Transcript sheet.
 * @param {string} nic - Student NIC number.
 * @return {{found: boolean, fullName?: string, batchNumber?: string, nts?: string, message?: string}}
 */
function lookupStudentByNic(nic) {
  if (!nic || !String(nic).trim()) {
    return { found: false, message: 'Student not found' };
  }
  const targetNic = String(nic).trim().toUpperCase();
  const sheet = getOrCreateSheet('Transcript');
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { found: false, message: 'Student not found' };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    const rowNic = String(data[i][1]).trim().toUpperCase();
    if (rowNic === targetNic) {
      return {
        found: true,
        fullName: data[i][0],
        batchNumber: data[i][2],
        nts: data[i][3]
      };
    }
  }

  return { found: false, message: 'Student not found' };
}

/**
 * Retrieves an existing record for editing from the specified sheet.
 * @param {string} sheetName - "Transcript", "Exam_Results", or "Medium_Letter".
 * @param {string} searchKey - NIC number or Index number to search.
 * @return {{found: boolean, record?: Object, message?: string}}
 */
function getRecordForEdit(sheetName, searchKey) {
  if (!searchKey || !String(searchKey).trim()) {
    return { found: false, message: 'Please enter a search key (NIC/Index Number).' };
  }

  const key = String(searchKey).trim().toUpperCase();
  const sheet = getOrCreateSheet(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { found: false, message: 'Record not found in ' + sheetName };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let match = false;

    if (sheetName === 'Transcript') {
      // Col 2 is NIC Number
      if (String(row[1]).trim().toUpperCase() === key) match = true;
    } else if (sheetName === 'Exam_Results') {
      // Col 3 is Index Number, Col 4 is NIC Number
      if (String(row[2]).trim().toUpperCase() === key || String(row[3]).trim().toUpperCase() === key) match = true;
    } else if (sheetName === 'Medium_Letter') {
      // Col 3 is NIC Number
      if (String(row[2]).trim().toUpperCase() === key) match = true;
    }

    if (match) {
      let record = {};
      if (sheetName === 'Transcript') {
        record = {
          fullName: row[0],
          nicNumber: row[1],
          batchNumber: row[2],
          nts: row[3],
          date: row[4]
        };
      } else if (sheetName === 'Exam_Results') {
        record = {
          title: row[0],
          fullName: row[1],
          indexNumber: row[2],
          nic: row[3],
          batchNumber: row[4],
          nts: row[5],
          hisHer: row[6],
          fromDate: row[7],
          toDate: row[8],
          paper1: row[9],
          paper2: row[10],
          paper3: row[11],
          paper4: row[12],
          paper5: row[13],
          practical: row[14],
          date: row[15]
        };
      } else if (sheetName === 'Medium_Letter') {
        record = {
          title: row[0],
          fullName: row[1],
          nic: row[2],
          batchNo: row[3],
          nts: row[4],
          fromDate: row[5],
          toDate: row[6],
          date: row[7]
        };
      }
      return { found: true, record: record, rowIndex: i + 2 };
    }
  }

  return { found: false, message: 'Record not found in ' + sheetName };
}

/**
 * Saves transcript / student entry data.
 * @param {Object} data - Transcript data.
 * @return {{success: boolean, message: string}}
 */
function saveTranscript(data) {
  if (!data) {
    return { success: false, message: 'Please fill all text fields' };
  }

  let fullName = data.fullName ? String(data.fullName).trim() : '';
  let nicNumber = data.nicNumber ? String(data.nicNumber).trim() : '';
  let batchNumber = data.batchNumber ? String(data.batchNumber).trim() : '';
  let nts = data.nts ? String(data.nts).trim() : '';
  let date = data.date ? String(data.date).trim() : '';

  if (!fullName || !nicNumber || !batchNumber || !nts || !date) {
    return { success: false, message: 'Please fill all text fields' };
  }

  fullName = properCase(fullName);
  nicNumber = nicNumber.toUpperCase();
  batchNumber = batchNumber.toUpperCase();

  const nicValidation = validateNIC(nicNumber);
  if (!nicValidation.valid) {
    return { success: false, message: nicValidation.message };
  }

  const sheet = getOrCreateSheet('Transcript');
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const nics = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < nics.length; i++) {
      if (String(nics[i][0]).trim().toUpperCase() === nicNumber) {
        return { success: false, message: 'This NIC number already exists. Please enter a unique NIC.' };
      }
    }
  }

  sheet.appendRow([fullName, nicNumber, batchNumber, nts, date]);
  return { success: true, message: 'Data saved successfully!' };
}

/**
 * Updates an existing transcript record.
 * @param {Object} data - Updated transcript data.
 * @return {{success: boolean, message: string}}
 */
function updateTranscript(data) {
  if (!data) return { success: false, message: 'Please fill all text fields' };

  let fullName = data.fullName ? String(data.fullName).trim() : '';
  let nicNumber = data.nicNumber ? String(data.nicNumber).trim() : '';
  let batchNumber = data.batchNumber ? String(data.batchNumber).trim() : '';
  let nts = data.nts ? String(data.nts).trim() : '';
  let date = data.date ? String(data.date).trim() : '';

  if (!fullName || !nicNumber || !batchNumber || !nts || !date) {
    return { success: false, message: 'Please fill all text fields' };
  }

  fullName = properCase(fullName);
  nicNumber = nicNumber.toUpperCase();
  batchNumber = batchNumber.toUpperCase();

  const nicValidation = validateNIC(nicNumber);
  if (!nicValidation.valid) return { success: false, message: nicValidation.message };

  const sheet = getOrCreateSheet('Transcript');
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return { success: false, message: 'Record not found to update.' };

  const dataVals = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < dataVals.length; i++) {
    if (String(dataVals[i][1]).trim().toUpperCase() === nicNumber) {
      sheet.getRange(i + 2, 1, 1, 5).setValues([[fullName, nicNumber, batchNumber, nts, date]]);
      return { success: true, message: 'Transcript record updated successfully!' };
    }
  }

  return { success: false, message: 'Student NIC not found in Transcript sheet.' };
}

/**
 * Saves exam results data.
 * @param {Object} data - Exam results data.
 * @return {{success: boolean, message: string}}
 */
function saveExamResults(data) {
  if (!data) {
    return { success: false, message: 'Please fill all text fields' };
  }

  const fields = [
    'title', 'fullName', 'indexNumber', 'nic', 'batchNumber',
    'nts', 'hisHer', 'fromDate', 'toDate',
    'paper1', 'paper2', 'paper3', 'paper4', 'practical', 'date'
  ];

  for (let field of fields) {
    if (!data[field] || !String(data[field]).trim()) {
      return { success: false, message: 'Please fill all text fields' };
    }
  }

  const sheet = getOrCreateSheet('Exam_Results');
  sheet.appendRow([
    data.title,
    data.fullName,
    data.indexNumber,
    data.nic,
    data.batchNumber,
    data.nts,
    data.hisHer,
    data.fromDate,
    data.toDate,
    data.paper1,
    data.paper2,
    data.paper3,
    data.paper4,
    data.paper5 || '',
    data.practical,
    data.date
  ]);

  return { success: true, message: 'Data saved successfully!' };
}

/**
 * Updates an existing exam results record.
 * @param {Object} data - Updated exam results data.
 * @return {{success: boolean, message: string}}
 */
function updateExamResults(data) {
  if (!data) return { success: false, message: 'Please fill all text fields' };

  const fields = [
    'title', 'fullName', 'indexNumber', 'nic', 'batchNumber',
    'nts', 'hisHer', 'fromDate', 'toDate',
    'paper1', 'paper2', 'paper3', 'paper4', 'practical', 'date'
  ];

  for (let field of fields) {
    if (!data[field] || !String(data[field]).trim()) {
      return { success: false, message: 'Please fill all text fields' };
    }
  }

  const sheet = getOrCreateSheet('Exam_Results');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: 'Record not found to update.' };

  const targetNic = String(data.nic).trim().toUpperCase();
  const targetIndex = String(data.indexNumber).trim().toUpperCase();
  const dataVals = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  for (let i = 0; i < dataVals.length; i++) {
    const rowNic = String(dataVals[i][3]).trim().toUpperCase();
    const rowIndex = String(dataVals[i][2]).trim().toUpperCase();

    if (rowNic === targetNic || rowIndex === targetIndex) {
      const updatedRow = [
        data.title,
        data.fullName,
        data.indexNumber,
        data.nic,
        data.batchNumber,
        data.nts,
        data.hisHer,
        data.fromDate,
        data.toDate,
        data.paper1,
        data.paper2,
        data.paper3,
        data.paper4,
        data.paper5 || '',
        data.practical,
        data.date
      ];
      sheet.getRange(i + 2, 1, 1, updatedRow.length).setValues([updatedRow]);
      return { success: true, message: 'Exam results record updated successfully!' };
    }
  }

  return { success: false, message: 'Exam record not found to update.' };
}

/**
 * Saves medium letter data.
 * @param {Object} data - Medium letter data.
 * @return {{success: boolean, message: string}}
 */
function saveMediumLetter(data) {
  if (!data) {
    return { success: false, message: 'Please fill all text fields' };
  }

  const fields = ['nic', 'title', 'fullName', 'batchNo', 'nts', 'fromDate', 'toDate', 'date'];
  for (let field of fields) {
    if (!data[field] || !String(data[field]).trim()) {
      return { success: false, message: 'Please fill all text fields' };
    }
  }

  const sheet = getOrCreateSheet('Medium_Letter');
  sheet.appendRow([
    data.title,
    data.fullName,
    data.nic,
    data.batchNo,
    data.nts,
    data.fromDate,
    data.toDate,
    data.date
  ]);

  return { success: true, message: 'Data saved successfully!' };
}

/**
 * Updates an existing medium letter record.
 * @param {Object} data - Updated medium letter data.
 * @return {{success: boolean, message: string}}
 */
function updateMediumLetter(data) {
  if (!data) return { success: false, message: 'Please fill all text fields' };

  const fields = ['nic', 'title', 'fullName', 'batchNo', 'nts', 'fromDate', 'toDate', 'date'];
  for (let field of fields) {
    if (!data[field] || !String(data[field]).trim()) {
      return { success: false, message: 'Please fill all text fields' };
    }
  }

  const sheet = getOrCreateSheet('Medium_Letter');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: 'Record not found to update.' };

  const targetNic = String(data.nic).trim().toUpperCase();
  const dataVals = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = 0; i < dataVals.length; i++) {
    const rowNic = String(dataVals[i][2]).trim().toUpperCase();
    if (rowNic === targetNic) {
      const updatedRow = [
        data.title,
        data.fullName,
        data.nic,
        data.batchNo,
        data.nts,
        data.fromDate,
        data.toDate,
        data.date
      ];
      sheet.getRange(i + 2, 1, 1, updatedRow.length).setValues([updatedRow]);
      return { success: true, message: 'Medium letter record updated successfully!' };
    }
  }

  return { success: false, message: 'Medium letter record not found to update.' };
}

/**
 * Serves the Home Page HTML interface when accessed as a Web App.
 * @param {Object} e - Event parameter.
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Management System - Home Page Access Tab')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Displays the main Home Page UI modal dialog in Google Sheets.
 */
function showHomePage() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(900)
    .setHeight(700)
    .setTitle('Home Page Access Tab - Student Management System');
  SpreadsheetApp.getUi().showModalDialog(html, 'Home Page Access Tab');
}

/**
 * Displays the Transcript Entry Form dialog.
 */
function showTranscriptForm() {
  showHomePage();
}

/**
 * Displays the Exam Results Form dialog.
 */
function showExamResultsForm() {
  showHomePage();
}

/**
 * Displays the Medium Letter Form dialog.
 */
function showMediumLetterForm() {
  showHomePage();
}

/**
 * Runs on spreadsheet open to add custom menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Student Management')
    .addItem('Home Page Access Tab', 'showHomePage')
    .addSeparator()
    .addItem('Setup Sheets & Headers', 'setupSpreadsheet')
    .addToUi();
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    authenticateUser,
    properCase,
    validateNIC,
    getNTSList,
    getTitleList,
    getHisHerList,
    lookupStudentByNic,
    getRecordForEdit,
    saveTranscript,
    updateTranscript,
    saveExamResults,
    updateExamResults,
    saveMediumLetter,
    updateMediumLetter,
    setupSpreadsheet,
    doGet,
    showHomePage,
    showTranscriptForm,
    showExamResultsForm,
    showMediumLetterForm,
    onOpen
  };
}
