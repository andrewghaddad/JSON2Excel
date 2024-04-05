# JSON2Excel: Node.js JSON to Excel Converter

This Node.js application takes in a JSON structure, parses it, and writes the data to an Excel spreadsheet.


## Prerequisites

Ensure you have the following installed on your machine:

- Node.js (https://nodejs.org)
- npm (Node Package Manager, comes with Node.js)
- excelJS


## Installation

1. Clone or download this repository to your local machine.

2. Navigate to the project directory in your terminal.

3. Install dependencies by running the following command:


```bash
npm install
npm install exceljs
```

## Installation

1. Ensure you have your JSON file ready that you want to convert to Excel.

2. Place your JSON file in the project directory.

3. Modify the input.json file to include your JSON data.

4. Run the following command to start the application:


```bash
node JSON2Excel.js
```

After the program executes successfully, you will find the generated Excel spreadsheet named `output.xlsx` in the project directory.


## Configuration

You can configure the behavior of the application by modifying the following parameters in the 

`index.js` file:
`inputFileName`: Name of the input JSON file.
`outputFileName`: Name of the output Excel file.
`sheetName`: Name of the sheet in the Excel file where data will be written.
`data`: JSON data structure to be parsed and written to the Excel file.


## Contributing

Contributions are welcome! If you find any bugs or have suggestions for improvement, please open an issue or create a pull request.