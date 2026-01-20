const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const SHEET_NAME = 'General Data';

async function inspectAddress() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        const sheet = workbook.getWorksheet(SHEET_NAME);
        if (!sheet) return;

        console.log('--- Inspecting Cols E (5) and L (12) for 2026_bidding ---');
        let count = 0;
        sheet.eachRow((row, i) => {
            if (i > 2 && count < 5) {
                const projectCodeVal = row.getCell(52).value;
                const projectCodeStr = projectCodeVal ? projectCodeVal.toString().toLowerCase() : '';

                if (projectCodeStr.includes('2026') && projectCodeStr.includes('bidding')) {
                    const eVal = row.getCell(5).value; // Shop Name / Address Snippet
                    const lVal = row.getCell(12).value; // Specific Address

                    console.log(`Row ${i}:`);
                    console.log(`  - E (Shop): ${eVal}`);
                    console.log(`  - L (Addr): ${lVal}`);
                    count++;
                }
            }
        });
    } catch (err) {
        console.error(err);
    }
}

inspectAddress();
