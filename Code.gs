/**
 * Creates a Google Form to collect staff information for examination-related duties.
 *
 * Fields collected:
 * - Employee Name
 * - Designation
 * - Current Workplace
 * - Telephone Number
 * - Preferred Location for Performing Examination Duties
 *
 * @return {GoogleAppsScript.Forms.Form} The created Google Form object.
 */
function createStaffExaminationForm() {
  const form = FormApp.create('Staff Examination Duty Form')
    .setTitle('Staff Information Form for Examination Duties')
    .setDescription('Please fill out this form to provide your details and preferences for examination-related duties.');

  form.addTextItem()
    .setTitle('Employee Name')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Designation')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Current Workplace')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Telephone Number')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Preferred Location for Performing Examination Duties')
    .setRequired(true);

  if (typeof Logger !== 'undefined') {
    Logger.log('Form created successfully.');
    Logger.log('Form Edit URL: ' + form.getEditUrl());
    Logger.log('Form Published URL: ' + form.getPublishedUrl());
  }

  return form;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { createStaffExaminationForm };
}
