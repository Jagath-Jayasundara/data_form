# Student Management System (Google Apps Script)

This project is a Google Apps Script project converted from VBA UserForms. It provides a full **Home Page Access Tab** and web-based interface for managing student transcripts, exam results, and medium of instruction letters in Google Sheets.

---

## Authentication & Login

The system features a built-in login form to secure access to student forms and records:
- **Default Username**: `admin`
- **Default Password**: `admin123`

When launched, users must authenticate before accessing the Home Page Access Tab or entry forms. A **Logout** button is provided in the upper header panel.

---

## Features & Modules

### 1. Home Page Access Tab
- Navigation dashboard offering single-click access to all user forms and sheet setups.
- Can be launched via the Spreadsheet Custom Menu (**Student Management > Home Page Access Tab**) or accessed as a Web App (`doGet`).

### 2. Transcript / Student Entry Form
- Fields: Full Name, NIC Number, Batch Number, NTS (Nursing Training School dropdown), Date.
- **Edit & Update Service**: Search by NIC Number using the "Load Record" search bar to edit and update existing transcript records.
- Automatic text formatting: ProperCase for Full Name, Uppercase for NIC and Batch Number.
- NIC Format Validation:
  - New NIC: Exactly 12 digits.
  - Old NIC: 11 characters (e.g. `912345678 V`).
- Duplicate NIC check against column B of the `Transcript` sheet.

### 3. Exam Results Entry Form
- Fields: Title, Full Name, Index Number, NIC Number, Batch Number, NTS, His/Her, From Date, To Date, Paper I, Paper II, Paper III, Paper IV, Paper V (conditional), Practical, Date.
- **Automatic Transcript Data Population**: Entering the student's NIC automatically retrieves and populates **Full Name**, **Batch Number**, and **NTS** from the `Transcript` data sheet.
- **Edit & Update Service**: Search by NIC or Index Number using the "Load Record" search bar to retrieve and update exam results.
- Results section arranged cleanly in 2 rows (Row 1: Paper I, II, III; Row 2: Paper IV, V, Practical).
- Dynamic conditional field: Enables **Paper V** if Batch Number year is $\ge 2018$.
- Saves records to the `Exam_Results` sheet.

### 4. Medium Letter Entry Form
- Fields: NIC Number, Title, Full Name, Batch Number, NTS, From Date, To Date, Date.
- **Automatic Transcript Data Population**: Entering the student's NIC automatically retrieves and populates **Full Name**, **Batch Number**, and **NTS** from the `Transcript` data sheet.
- **Edit & Update Service**: Search by NIC Number to load and update existing medium letter records.
- Saves records to the `Medium_Letter` sheet.

---

## Sheets Structure

When running **Student Management > Setup Sheets & Headers**:
- **`Home`**: Overview sheet with system module descriptions and links.
- **`Transcript`**: Stores registered student details (`[Full Name, NIC Number, Batch Number, NTS, Date]`).
- **`Exam_Results`**: Stores examination paper results (`[Title, Full Name, Index Number, NIC, Batch Number, NTS, His/Her, From Date, To Date, Paper I, Paper II, Paper III, Paper IV, Paper V, Practical, Date]`).
- **`Medium_Letter`**: Stores medium of instruction letter records (`[Title, Full Name, NIC, Batch Number, NTS, From Date, To Date, Date]`).

---

## How to Set Up in Google Apps Script

### Option 1: Copying Code to Google Sheets Apps Script Editor
1. Open your Google Sheet.
2. Click **Extensions** > **Apps Script**.
3. Replace the contents of `Code.gs` with the code in `Code.gs`.
4. Create an HTML file named `Index.html` and paste the contents of `Index.html`.
5. Save the project (`Ctrl+S` / `Cmd+S`).
6. Refresh your Google Sheet. A new menu named **Student Management** will appear on the menu bar.
7. Click **Student Management** > **Setup Sheets & Headers** to create all target sheets.
8. Click **Student Management** > **Home Page Access Tab** to launch the interface.

---

## Running Unit Tests Locally

To run the Node.js native unit tests locally:

```bash
npm test
```
