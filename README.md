
# Registration Device Growth Report

This project generates a report showing month-wise registration growth per device type, grouped by Calley Status, using Node.js and Express. It processes an Excel file (`RegistrationReport.xlsx`) and provides the report in multiple formats.

## Setup
1. Clone the repo:
   git clone task3
   cd task3
2. Install dependencies:
   npm install
3. Ensure `RegistrationReport.xlsx` is in the project folder with columns: `Registration Date` (e.g., "Mar 21 2025 10:19AM"), `RegistrationDevice`, and `Calley Status`.

## How to Use
Run the project with:

node index.js

The server will start on `http://localhost:3000`. You can access the report in three ways:

### 1. Manual Report in Folder
- An Excel file (`registration_report.xlsx`) is created in the project folder.
- Open it to see the report with columns: `Device`, `Calley Status`, monthly counts, and `Total`.

### 2. Create JSON Data
- The report data is converted to JSON format internally.
- To see the JSON, add `console.log(reportJson)` in `app.js` after the `reportJson` assignment and rerun.

### 3. Send JSON via Routes
- Access the JSON data via an API:
  - Visit `http://localhost:3000/report` in a browser to see the JSON.
- Download the Excel file:
  - Visit `http://localhost:3000/download-report` to download `registration_report.xlsx`.

## Notes
- If `RegistrationReport.xlsx` is missing or has invalid data, check the server logs for errors.
