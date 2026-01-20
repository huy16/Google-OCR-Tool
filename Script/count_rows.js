const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const SHEET_NAME = 'View and Take Data';

async function countRows() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        const sheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];

        if (!sheet) {
            console.log('Sheet not found');
            return;
        }

        console.log(`Sheet: ${sheet.name}`);
        console.log(`Total Rows (rowCount property): ${sheet.rowCount}`);

        // Count actual data rows (assuming row 1 is title, 2 is header)
        let dataCount = 0;
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber > 2 && row.hasValues) {
                // Check if meaningful data exists (e.g. column 5 "BHX_..." or column 1 "StoreId" if exists)
                // Based on previous script, Address is around col 5
                const cellVal = row.getCell(5).value;
                if (cellVal) dataCount++;
            }
        });
        console.log(`Estimated Data Rows (Row > 2 with value in Col 5): ${dataCount}`);

    } catch (err) {
        console.error('Error:', err.message);
    }
}

countRows();
