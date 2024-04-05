const fs = require('fs');
const ExcelJS = require('exceljs');

// Configuration parameters
const inputFileName = 'input.json'; // Name of the input JSON file
const date = new Date().toDateString(); // Timestamp of output Excel file
const outputFilePath = './node/files/'; // File path where output Excel file will be written to
const outputFileName = outputFilePath + 'output ' + date + '.xlsx'; // Name of the output Excel file
const sheetName = 'News For ' + date; // Name of the sheet in the Excel file
const data = require(`../${inputFileName}`); // JSON data structure retrieved from input JSON file
const dataObject = data['articles']; // JSON data structure to be parsed and written to the Excel file
const headersObject = data['articles'][0]; // JSON data structure of the headers to be written to the Excel file

// Function to convert JSON data to Excel
async function convertJsonToExcel() {
    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(sheetName);

    // Write headers
    const headers = Object.keys(headersObject);
    worksheet.addRow(headers);

    // Write data
    dataObject.forEach(obj => {
        const row = [];
        headers.forEach(header => {
            row.push(obj[header]);
        });
        worksheet.addRow(row);
    });

    // Save the workbook
    await workbook.xlsx.writeFile(outputFileName);
    console.log(`Excel file "${outputFileName}" generated successfully.`);
}

// Main function
async function main() {
    try {
        await convertJsonToExcel();
    } catch (error) {
        console.error('[JSON2Excel] Error:', error.message);
    }
}

// Execute the main function
main();