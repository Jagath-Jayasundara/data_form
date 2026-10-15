const test = require('node:test');
const assert = require('node:assert');

class MockTextItem {
  constructor() {
    this.title = '';
    this.required = false;
  }
  setTitle(title) {
    this.title = title;
    return this;
  }
  setRequired(required) {
    this.required = required;
    return this;
  }
}

class MockForm {
  constructor(title) {
    this.name = title;
    this.title = '';
    this.description = '';
    this.items = [];
  }
  setTitle(title) {
    this.title = title;
    return this;
  }
  setDescription(description) {
    this.description = description;
    return this;
  }
  addTextItem() {
    const item = new MockTextItem();
    this.items.push(item);
    return item;
  }
  getEditUrl() {
    return 'https://docs.google.com/forms/d/mock-id/edit';
  }
  getPublishedUrl() {
    return 'https://docs.google.com/forms/d/e/mock-id/viewform';
  }
}

global.FormApp = {
  create: (title) => new MockForm(title)
};

const { createStaffExaminationForm } = require('../Code.gs');

test('createStaffExaminationForm creates form with correct fields', () => {
  const form = createStaffExaminationForm();

  assert.strictEqual(form.name, 'Staff Examination Duty Form');
  assert.strictEqual(form.title, 'Staff Information Form for Examination Duties');
  assert.strictEqual(
    form.description,
    'Please fill out this form to provide your details and preferences for examination-related duties.'
  );

  assert.strictEqual(form.items.length, 5);

  const expectedFields = [
    'Employee Name',
    'Designation',
    'Current Workplace',
    'Telephone Number',
    'Preferred Location for Performing Examination Duties'
  ];

  expectedFields.forEach((expectedTitle, index) => {
    assert.strictEqual(form.items[index].title, expectedTitle);
    assert.strictEqual(form.items[index].required, true);
  });
});
