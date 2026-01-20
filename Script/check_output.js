const ExcelJS = require('exceljs');
const path = require('path');

async function checkOutput() {
    const wb = new ExcelJS.Workbook();
    const filePath = path.join(__dirname, '../1. Database/Geocode_Results_Simple_v2.xlsx');

    try {
        await wb.xlsx.readFile(filePath);

        console.log('Worksheets found:', wb.worksheets.map(ws => ws.name).join(', '));

        const sheet = wb.getWorksheet('Results');
        if (sheet) {
            console.log(`Sheet 'Results' found. Rows: ${sheet.rowCount}`);
            console.log('Headers:', sheet.getRow(1).values.slice(1).join(' | '));
            if (sheet.rowCount > 1) {
                console.log('Last Row:', sheet.getRow(sheet.rowCount).values.slice(1).join(' | '));
            }
        } else {
            console.log('Sheet "Results" NOT found!');
        }

    } catch (err) {
        console.error('Error:', err.message);
    }
}

checkOutput();
