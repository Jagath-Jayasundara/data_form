# Staff Examination Duty Form (Google Apps Script)

This project contains Google Apps Script code to automatically create a Google Form for collecting staff examination duty details.

## Fields Collected
1. Employee Name
2. Designation
3. Current Workplace
4. Telephone Number
5. Preferred Location for Performing Examination Duties

---

## How to Run in Google Apps Script Editor

### Option 1: Copy Code Directly to Google Apps Script Editor
1. Go to [script.google.com](https://script.google.com/) and click **New project**.
2. Replace the contents of `Code.gs` with the code in `Code.gs` from this project.
3. Click the **Save** icon (or press `Ctrl+S` / `Cmd+S`).
4. At the top editor bar, select the function **`createStaffExaminationForm`** from the dropdown menu.
5. Click **Run**.
6. When running for the first time, Google Apps Script will request permissions to manage your Google Forms:
   - Click **Review permissions**.
   - Choose your Google account.
   - Click **Advanced** > **Go to Untitled project (unsafe)** (or your project name).
   - Click **Allow**.
7. View the **Execution log** at the bottom of the Apps Script editor to find the generated Form Edit URL and Form Published URL.

---

### Option 2: Using `clasp` (Command Line)
If you use Google Apps Script Command Line Tool (`clasp`):
1. Install clasp: `npm install -g @google/clasp`
2. Log in to your Google account: `clasp login`
3. Push files to your project: `clasp push`
4. Open the editor: `clasp open` and run `createStaffExaminationForm`.

---

## Running Unit Tests Locally

To run unit tests locally:
```bash
npm test
```
