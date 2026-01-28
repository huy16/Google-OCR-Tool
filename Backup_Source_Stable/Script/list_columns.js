const ExcelJS = require('exceljs');
const path = require('path');

(async () => {
    const filePath = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('General Data') || workbook.worksheets[0];

    // Assuming Header is Row 2 based on previous knowledge
    const headerRow = sheet.getRow(2);
    console.log('--- Column Headers (Row 2) ---');
    headerRow.eachCell((cell, colNum) => {
        const val = cell.value ? cell.value.toString().trim() : '';
        if (val) {
            console.log(`Col ${colNum}: ${val}`);
        }
    });

    // Also check Row 1 just in case
    const row1 = sheet.getRow(1);
    console.log('\n--- Row 1 Preview ---');
    row1.eachCell((cell, colNum) => {
        const val = cell.value ? cell.value.toString().trim() : '';
        if (val) console.log(`Col ${colNum}: ${val}`);
    });

})();
