const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Serve the HTML form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Handle form submission
app.post('/submit', (req, res) => {
  const { name, company, email, phone_number } = req.body;

  // Define the path of the Excel file
  const filePath = path.join(__dirname, 'contacts.xlsx');
  
  let workbook;

  // Check if the Excel file already exists
  if (fs.existsSync(filePath)) {
    // Read the existing workbook
    workbook = XLSX.readFile(filePath);
  } else {
    // If file doesn't exist, create a new workbook and initialize with headers
    workbook = XLSX.utils.book_new();
    const headers = [['Name', 'Company', 'Email', 'Phone Number']];
    const sheet = XLSX.utils.aoa_to_sheet(headers); // Create a sheet with headers
    workbook.Sheets['Contacts'] = sheet; // Assign the sheet to the workbook
    XLSX.utils.book_append_sheet(workbook, sheet, 'Contacts');
  }

  // Get the existing sheet or initialize if it doesn't exist
  const sheet = workbook.Sheets['Contacts'];

  // Convert the existing sheet into a JSON array, including header row
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }) || [];

  // Add the new row with user input data
  data.push([name, company, email, phone_number]);

  // Create a new sheet with updated data (including headers)
  const updatedSheet = XLSX.utils.aoa_to_sheet(data);

  // Update the workbook with the new sheet
  workbook.Sheets['Contacts'] = updatedSheet;

  // Save the updated workbook back to the file
  XLSX.writeFile(workbook, filePath);

  res.send('Data saved successfully! <a href="/">Go back to the form</a>');
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
