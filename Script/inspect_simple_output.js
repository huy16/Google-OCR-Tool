const ExcelJS = require('exceljs');
const path = require('path');

const OUTPUT_FILE_PATH = path.join(__dirname, '../1. Database/Geocode_Results_Simple.xlsx');

async function inspectSimple() {
    const workbook = new ExcelJS.Workbook();
    if (path.extname(OUTPUT_FILE_PATH) !== '.xlsx') {
        console.log("File is not xlsx");
        return;
    }

    try {
        await workbook.xlsx.readFile(OUTPUT_FILE_PATH);
        const sheet = workbook.getWorksheet('Results');
        if (!sheet) {
            console.log("Sheet 'Results' not found.");
            // Print all sheet names
            workbook.eachSheet(s => console.log('Sheet found:', s.name));
            return;
        }

        console.log(`Sheet 'Results' found. Row Count: ${sheet.rowCount}`);
        sheet.eachRow((row, i) => {
            console.log(`Row ${i}: ${JSON.stringify(row.values)}`);
        });

    } catch (err) {
        console.error("Error reading file:", err);
    }
}

inspectSimple();
