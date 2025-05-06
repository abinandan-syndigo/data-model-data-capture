const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const config = require('./config');

// Define input and output directories
const inputFolderPath = './data/input';
const outputFolderPath = './data/output';

// Ensure the output folder exists
if (!fs.existsSync(outputFolderPath)) {
    fs.mkdirSync(outputFolderPath, { recursive: true });
}

// Read the files in the input folder
const files = fs.readdirSync(inputFolderPath).filter(file => file.endsWith('.xlsx'));

// Create a map to store unique values and their counts per file
const dataTypeCountMap = {};

// Process each file in the input folder
files.forEach(file => {
    const filePath = path.join(inputFolderPath, file);
    console.log(`Processing file: ${file}`);

    try {
        // Read the Excel file
        const workbook = XLSX.readFile(filePath);
        
        // Get the "ATTRIBUTES" sheet, if it exists
        const sheet = workbook.Sheets["ATTRIBUTES"];
        if (!sheet) {
            console.warn(`Sheet "ATTRIBUTES" not found in ${file}, skipping.`);
            return;
        }
        
        // Convert the sheet to JSON format
        const data = XLSX.utils.sheet_to_json(sheet);
        console.log(`Found ${data.length} rows of data in ${file}`);

        // Initialize a new map for the current file
        const fileDataTypeCountMap = Object.create(null);

        // Iterate over the rows of data
        data.forEach(row => {
            const dataType = row['DATA TYPE'];
            if (!dataType) return;
            fileDataTypeCountMap[dataType] = (fileDataTypeCountMap[dataType] || 0) + 1;
        });

        console.log(`Found ${Object.keys(fileDataTypeCountMap).length} unique data types in ${file}`);

        // Store the counts for the current file in the main map
        for (const [key, value] of Object.entries(fileDataTypeCountMap)) {
            if (!dataTypeCountMap[key]) {
                dataTypeCountMap[key] = {};
            }
            dataTypeCountMap[key][file] = value;
        }
    } catch (error) {
        console.error(`Error processing file ${file}:`, error.message);
    }
});

// Prepare the data for writing to Excel
const outputData = [];
const headers = ['DATA TYPE', ...files];
outputData.push(headers);

for (const [dataType, counts] of Object.entries(dataTypeCountMap)) {
    const row = [dataType];
    files.forEach(file => {
        row.push(counts[file] || 0);
    });
    outputData.push(row);
}

// Create a new workbook and add a sheet with the output data
const outputWorkbook = XLSX.utils.book_new();
const outputSheet = XLSX.utils.aoa_to_sheet(outputData);
XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'Output');

// Generate a timestamped output file name
const timestamp = new Date().toISOString().replace(/[-:]/g, '').split('.')[0];
const outputFileName = `output_${timestamp}.xlsx`;
const outputFilePath = path.join(outputFolderPath, outputFileName);

// Write the output workbook to the specified output folder
XLSX.writeFile(outputWorkbook, outputFilePath);

console.log(`✅ Output written to ${outputFilePath}`);