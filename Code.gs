
const SPREADSHEET_PROPERTY = 'EXAM_STAFF_SPREADSHEET_ID';

// Default general capacities per location
const LOCATION_CAPACITIES = {
  'PBCN-Colombo': 9,
  'CON-Ampara': 3,
  'CON-Anuradhapura': 2,
  'CON-Galle': 3,
  'CON-Jaffna': 2,
  'CON-Kandy': 3,
  'CON-Kurunegala': 3,
  'CON-Ratnapura': 2,
};

// Default quota allocated specifically for staff in the Junior group (2 vacancies per location)
const JUNIOR_LOCATION_CAPACITIES = {
  'PBCN-Colombo': 2,
  'CON-Ampara': 2,
  'CON-Anuradhapura': 2,
  'CON-Galle': 2,
  'CON-Jaffna': 2,
  'CON-Kandy': 2,
  'CON-Kurunegala': 2,
  'CON-Ratnapura': 2,
};

function setup() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty(SPREADSHEET_PROPERTY);
  let spreadsheet;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create(
      'Examination Staff Information Responses'
    );
    properties.setProperty(SPREADSHEET_PROPERTY, spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName('Responses');

  if (!sheet) {
    sheet = spreadsheet.getSheets()[0];
    sheet.setName('Responses');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Submission Date',
      'Employee Name',
      'NIC Number',
      'Designation',
      'Category',
      'Current Workplace',
      'Telephone Number',
      'Preferred Examination Duty Location'
    ]);

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('C:C').setNumberFormat('@');
    sheet.getRange('G:G').setNumberFormat('@');
    sheet.autoResizeColumns(1, 8);
  }

  Logger.log('Response spreadsheet: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Examination Staff Information Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFormOptions(category) {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_PROPERTY);

  if (!spreadsheetId) {
    throw new Error(
      'The form has not been initialized. Please contact the administrator.'
    );
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('Responses');

  if (!sheet) {
    throw new Error('The response sheet could not be found.');
  }

  const lastRow = sheet.getLastRow();
  const isJunior = category === 'Junior';
  const capacities = isJunior ? JUNIOR_LOCATION_CAPACITIES : LOCATION_CAPACITIES;

  let submittedRecords = [];
  if (lastRow >= 2) {
    // Columns: Category (Col 5), Preferred Location (Col 8)
    const values = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    submittedRecords = values.map(row => ({
      category: (row[4] || '').trim(),
      location: (row[7] || '').trim()
    }));
  }

  return {
    category: isJunior ? 'Junior' : 'General',
    locations: Object.entries(capacities).map(
      ([name, capacity]) => {
        let registered = 0;
        if (isJunior) {
          registered = submittedRecords.filter(
            rec => rec.location === name && rec.category === 'Junior'
          ).length;
        } else {
          registered = submittedRecords.filter(
            rec => rec.location === name
          ).length;
        }

        return {
          name: name,
          capacity: capacity,
          registered: registered,
          remaining: Math.max(0, capacity - registered),
          full: registered >= capacity
        };
      }
    )
  };
}

function getEmployeeCategory(designation) {
  if (!designation || typeof designation !== 'string') {
    return 'Senior';
  }

  const clean = designation.trim().toUpperCase();

  // Junior group consists of positions designated as S.K.S. and K.K.S.
  const isJunior = clean.includes('S.K.S') || clean.includes('SKS') ||
                   clean.includes('K.K.S') || clean.includes('KKS');

  if (isJunior) {
    return 'Junior';
  }

  return 'Senior';
}

function submitForm(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Invalid form submission. Please try again.');
  }

  const employeeName = cleanField(data.employeeName, 'Employee Name', 150);
  const nic = cleanField(data.nic, 'NIC Number', 12);
  const designation = cleanField(data.designation, 'Designation', 150);
  const workplace = cleanField(data.workplace, 'Current Workplace', 200);
  const telephone = cleanField(data.telephone, 'Telephone Number', 30);
  const preferredLocation = cleanField(
    data.preferredLocation,
    'Preferred Location',
    100
  );

  const normalizedNic = validateAndNormalizeNic(nic);
  const category = getEmployeeCategory(designation);
  const isJunior = category === 'Junior';
  const targetCapacities = isJunior ? JUNIOR_LOCATION_CAPACITIES : LOCATION_CAPACITIES;

  if (!Object.prototype.hasOwnProperty.call(targetCapacities, preferredLocation)) {
    throw new Error('Please select a valid preferred location.');
  }

  if (!/^[0-9+()\-\s]{7,30}$/.test(telephone)) {
    throw new Error(
      'Enter a valid telephone number using digits and optional +, spaces, brackets or hyphens.'
    );
  }

  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_PROPERTY);

  if (!spreadsheetId) {
    throw new Error(
      'The form has not been initialized. Please contact the administrator.'
    );
  }

  // Keep the duplicate check, capacity check, and row insertion under the same lock.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('Responses');

    if (!sheet) {
      throw new Error('The response sheet could not be found.');
    }

    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const existingValues = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
      const normalizedSubmittedPhone = normalizePhoneNumber(telephone);

      for (let i = 0; i < existingValues.length; i++) {
        const existingNic = (existingValues[i][2] || '').trim().toUpperCase();
        const existingPhone = normalizePhoneNumber(existingValues[i][6] || '');

        if (existingNic === normalizedNic.toUpperCase()) {
          throw new Error('An entry with this NIC number has already been submitted.');
        }

        if (existingPhone && normalizedSubmittedPhone && existingPhone === normalizedSubmittedPhone) {
          throw new Error('An entry with this telephone number has already been submitted.');
        }
      }

      // Check capacity for the specific group
      const capacity = targetCapacities[preferredLocation];
      let registered = 0;

      for (let i = 0; i < existingValues.length; i++) {
        const rowCategory = (existingValues[i][4] || '').trim();
        const rowLocation = (existingValues[i][7] || '').trim();

        if (rowLocation === preferredLocation) {
          if (isJunior) {
            if (rowCategory === 'Junior') {
              registered++;
            }
          } else {
            registered++;
          }
        }
      }

      if (registered >= capacity) {
        throw new Error(
          'Sorry, ' + preferredLocation +
          ' has reached its maximum capacity of ' + capacity +
          (isJunior ? ' staff members for the Junior quota.' : ' staff members.') +
          ' Please select another location.'
        );
      }
    }

    sheet.appendRow([
      new Date(),
      employeeName,
      normalizedNic,
      designation,
      category,
      workplace,
      "'" + telephone,
      preferredLocation
    ]);

    return {
      success: true,
      message: 'Your information has been submitted successfully.'
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizePhoneNumber(phone) {
  if (!phone || typeof phone !== 'string') return '';
  return phone.replace(/[^0-9]/g, '');
}

function validateAndNormalizeNic(nic) {
  const isOldFormat = /^[0-9]{9}[vV]$/.test(nic);
  const isNewFormat = /^[0-9]{12}$/.test(nic);

  if (!isOldFormat && !isNewFormat) {
    throw new Error(
      'Please enter a valid NIC number. Old format requires 9 digits followed by "V", and new format requires 12 digits.'
    );
  }

  return isOldFormat ? nic.slice(0, 9) + 'V' : nic;
}

function cleanField(value, fieldName, maxLength) {
  if (typeof value !== 'string') {
    throw new Error('Please enter ' + fieldName + '.');
  }

  const cleaned = value.trim();

  if (!cleaned) {
    throw new Error(fieldName + ' is required.');
  }

  if (cleaned.length > maxLength) {
    throw new Error(fieldName + ' is too long.');
  }

  return cleaned;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    SPREADSHEET_PROPERTY,
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
  };
}
