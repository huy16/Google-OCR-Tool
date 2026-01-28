const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');

async function listSheets() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        console.log('Sheets found:');
        workbook.worksheets.forEach((sheet, index) => {
            console.log(`${index}: ${sheet.name}`);
        });
    } catch (err) {
        console.error('Error:', err.message);
    }
}

listSheets();
