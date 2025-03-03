const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const inputFolderPath = './data/input';
const files = fs.readdirSync(inputFolderPath);

if (files.length === 0) {
    console.error('No files found in the input folder.');
    process.exit(1);
}

const firstFile = files[0];
const firstFilePath = path.join(inputFolderPath, firstFile);

const workbook = XLSX.readFile(firstFilePath);
const sheet = workbook.Sheets["ATTRIBUTES"];

// Convert the sheet to JSON format
const data = XLSX.utils.sheet_to_json(sheet);

// Create a map to store unique values and their counts
const dataTypeCountMap = {};

// Iterate over the rows of data
data.forEach(row => {
    const dataType = row["DATA TYPE"];
    if (dataTypeCountMap[dataType]) {
        dataTypeCountMap[dataType]++;
    } else {
        dataTypeCountMap[dataType] = 1;
    }
});

// Convert the map to an array of objects for writing to Excel
const outputData = Object.keys(dataTypeCountMap).map(key => ({
    DataTypes: key,
    Count: dataTypeCountMap[key]
}));

// Create a new workbook and add a sheet with the output data
const outputWorkbook = XLSX.utils.book_new();
const outputSheet = XLSX.utils.json_to_sheet(outputData);
XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'Output');

// Generate a timestamp
const timestamp = new Date().toISOString().replace(/[-:]/g, '').split('.')[0];

// Write the output workbook to a new Excel file named with the timestamp
const outputFileName = `output_${timestamp}.xlsx`;
XLSX.writeFile(outputWorkbook, outputFileName);

console.log(`Output written to ${outputFileName}`);
