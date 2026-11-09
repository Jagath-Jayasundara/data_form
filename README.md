# Examination Staff Information System (Google Apps Script)

This project contains a Google Apps Script Web Application and Spreadsheet backend to collect information from staff members for examination-related duties with location capacity management.

## Information Collected
1. Employee Name
2. NIC Number (Validates 9 digits followed by 'V' for old format, or 12 digits for new format)
3. Designation
4. Current Workplace
5. Telephone Number
6. Preferred Location for Performing Examination Duties

---

## How to Set Up and Run

### 1. Create a Google Apps Script Project
1. Go to [script.google.com](https://script.google.com/) and click **New project**.
2. Name the project (e.g. `Examination Staff Information System`).
3. Replace the contents of `Code.gs` with the code in `Code.gs`.
4. In the Apps Script editor, click the **+** (Add a file) button next to "Files", select **HTML**, name it `Index`, and paste the contents of `Index.html`.
5. (Optional) In Project Settings, enable "Show 'appsscript.json' manifest file in editor" and copy `appsscript.json` if configuring OAuth scopes.

### 2. Initialize the Response Spreadsheet
1. In the Apps Script toolbar, select the **`setup`** function from the dropdown.
2. Click **Run**.
3. Authorize the script when prompted.
4. Check the **Execution log** to see the URL of the automatically created Google Sheet titled `Examination Staff Information Responses`.

### 3. Deploy as a Web Application
1. Click **Deploy** > **New deployment** (top right).
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set the following configuration:
   - **Description**: Examination Staff Duty Registration
   - **Execute as**: Me (your Google account)
   - **Who has access**: Anyone (or Anyone within your organization)
4. Click **Deploy**.
5. Copy the generated **Web App URL** and share it with staff members.

---

## Testing Locally
Run unit tests with:
```bash
npm test
```
