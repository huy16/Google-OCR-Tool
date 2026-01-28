const ExcelJS = require('exceljs');
const path = require('path');

const OUTPUT_FILE_PATH = path.join(__dirname, '../1. Database/Output_Geocode_Results.xlsx');
const SHEET_NAME = 'General Data';

async function inspectOutput() {
    const workbook = new ExcelJS.Workbook();
    try {
        console.log(`Reading output file: ${OUTPUT_FILE_PATH}`);
        await workbook.xlsx.readFile(OUTPUT_FILE_PATH);

        const sheet = workbook.getWorksheet(SHEET_NAME);
        if (!sheet) {
            console.log(`Sheet '${SHEET_NAME}' NOT found in output.`);
            return;
        }

        console.log(`\nFound Sheet: '${sheet.name}'`);

        // Rows we know were processed based on logs: 18, 362, 368, 369, 370, 372
        const checkRows = [18, 50, 100, 362, 368, 369, 370, 372];
        const linkCol = 11; // J=10, K=11. Script mapped Link to 11 (Google Map)

        console.log('--- Inspecting Processed Rows ---');
        checkRows.forEach(r => {
            const row = sheet.getRow(r);
            const linkVal = row.getCell(linkCol).value;
            const addressVal = row.getCell(5).value; // Shop Name

            console.log(`Row ${r}:`);
            console.log(`  - Shop: ${addressVal}`);
            console.log(`  - Link (Col 11): ${JSON.stringify(linkVal)}`);
        });

    } catch (err) {
        console.error('Error:', err);
    }
}

inspectOutput();
