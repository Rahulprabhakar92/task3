const express = require('express');
const XLSX = require('xlsx');
const moment = require('moment');
const path = require('path'); // Added for file path handling

const app = express();
const port = 3000;




function generateReport() {

  const workbook = XLSX.readFile('RegistrationReport .xlsx'); 
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet);

  // Process data into nested structure
  const reportData = {};
  rows.forEach((row) => {
    const date = moment(row['Registration Date'], 'MMM DD YYYY hh:mmA');
    // Skip rows with invalid dates
    if (!date.isValid()) {
      console.warn(`Skipping row with invalid date: ${row['Registration Date']}`);
      return;
    }
    const month = date.format('MMM YYYY');

    const device = row['RegistrationDevice'].toLowerCase() === 'ios' ? 'iOS' : row['RegistrationDevice'];
    const status = row['Calley Status'];

    if (!reportData[device]) reportData[device] = {};
    if (!reportData[device][status]) reportData[device][status] = {};
    if (!reportData[device][status][month]) reportData[device][status][month] = 0;

    reportData[device][status][month]++;
  });

  // Get unique months dynamically and sort chronologically
  const months = [...new Set(rows
    .map(row => moment(row['Registration Date'], 'MMM DD YYYY hh:mmA'))
    .filter(date => date.isValid())
    .map(date => date.format('MMM YYYY')))]
    .sort((a, b) => moment(a, 'MMM YYYY').valueOf() - moment(b, 'MMM YYYY').valueOf());

  // Build the table data
  const tableData = [];
  
  // Add header row
  const header = ['Device', 'Calley Status', ...months, 'Total'];
  tableData.push(header);

  // Add data rows
  for (const device in reportData) {
    for (const status in reportData[device]) {
      const row = [device, status];
      let total = 0;
      months.forEach(month => {
        const count = reportData[device][status][month] || 0;
        row.push(count);
        total += count;
      });
      row.push(total);
      tableData.push(row);
    }
  }

  // Create a new workbook and sheet
  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.aoa_to_sheet(tableData);
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Report');

  // Write the report to a file
  XLSX.writeFile(newWorkbook, 'registration_report.xlsx');

  console.log('Report generated: registration_report.xlsx');

  // Convert tableData to JSON for the API
  reportJson = tableData.slice(1).map(row => {
    const obj = {};
    header.forEach((key, index) => {
      obj[key] = row[index];
    });
    return obj;
  });

  return reportJson; // Return for better error handling
}

try {
  generateReport();

  // API endpoint to send report data as JSON
  app.get('/report', (req, res) => {
    if (!reportJson || reportJson.length === 0) {
      return res.status(500).json({ error: 'Report data not available' });
    }
    res.json(reportJson);
  });

  // Endpoint to download the Excel file
  app.get('/download-report', (req, res) => {
    const file = path.join(__dirname, 'registration_report.xlsx');
    res.download(file, 'registrationnew_report.xlsx', (err) => {
      if (err) {
        console.error('Error downloading file:', err);
        res.status(500).json({ error: 'Failed to download report' });
      }
    });
  });

  app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
  });
} catch (err) {
  console.error('Error generating report:', err);
  process.exit(1); // Exit if the initial report generation fails
}