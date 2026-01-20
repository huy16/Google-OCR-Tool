const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const SHEET_NAME = 'General Data';

async function countRows() {
    const workbook = new ExcelJS.Workbook();
    try {
        console.log(`Reading file: ${EXCEL_FILE_PATH}`);
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);

        const sheet = workbook.getWorksheet(SHEET_NAME);
        if (!sheet) {
            console.log(`Sheet '${SHEET_NAME}' NOT found.`);
            return;
        }

        console.log(`\nFound Sheet: '${sheet.name}'`);

        // Count Rows for "2026_bidding" in Col 52
        const projectCol = 52; // "MÃ DỰ ÁN" is Col 52 based on debug output
        const sttCol = 1;

        console.log(`\n--- Counting Rows (Val '2026_bidding' in Col ${projectCol}) ---`);
        let matchCount = 0;
        let totalSttCount = 0;

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber > 2) {
                const sttVal = row.getCell(sttCol).value;
                if (sttVal) totalSttCount++;

                const projectVal = row.getCell(projectCol).value;
                const projectStr = projectVal ? projectVal.toString().toLowerCase() : '';

                // Match "2026_bidding"
                // User said: "lọc dự án 2026_bidding thôi"
                if (projectStr.includes('2026') && projectStr.includes('bidding')) {
                    matchCount++;
                }
            }
        });

        console.log(`Total rows with data in STT: ${totalSttCount}`);
        console.log(`Total rows matching '2026_bidding' (Col ${projectCol}): ${matchCount}`);

    } catch (err) {
        console.error('Error:', err);
    }
}

countRows();
