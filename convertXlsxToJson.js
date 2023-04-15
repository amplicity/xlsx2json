// Import required libraries
const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');

// Function to convert XLSX to JSON
function xlsxToJson(inputFile, outputFile) {
  try {
    // Read the XLSX file
    const workbook = XLSX.readFile(inputFile, { cellDates: true });

    // Get the first sheet name
    const sheetName = workbook.SheetNames[0];

    // Convert sheet to JSON
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { dateNF: 'DD/MM/YY h:mm A' });

    // Save JSON to file
    fs.writeFileSync(outputFile, JSON.stringify(jsonData, null, 2));
    console.log(`Successfully converted '${inputFile}' to '${outputFile}'`);
  } catch (error) {
    console.error(`Error converting file: ${error}`);
  }
}

// Set up readline interface for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Ask for input file path
rl.question('Please enter the path to the input.xlsx file: ', (inputFile) => {
  // Define output file path
  const outputFile = './output.json';

  // Convert the XLSX file to JSON
  xlsxToJson(inputFile, outputFile);

  // Close the readline interface
  rl.close();
});
