const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const SHEET_NAME = 'General Data';
const HEADER_ROW = 2;

async function inspectProvinceColumn() {
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
        console.log(`Total rows in sheet: ${sheet.rowCount}`);

        // First, let's check header row to find column B
        console.log('\n--- Inspecting Header Row (Row 2) ---');
        const headerRow = sheet.getRow(HEADER_ROW);

        // Print first 10 columns to see headers
        for (let col = 1; col <= 10; col++) {
            const cellVal = headerRow.getCell(col).value;
            console.log(`  Col ${col}: ${cellVal}`);
        }

        // Now let's look at Column B (Col 2) specifically
        const provinceCol = 2; // Column B
        console.log(`\n--- Inspecting Column B (Col ${provinceCol}) - Expected: "Tỉnh" ---`);

        // Check first few data rows
        console.log('\nSample data from Column B (Rows 3-10):');
        for (let r = 3; r <= 10; r++) {
            const row = sheet.getRow(r);
            const provinceVal = row.getCell(provinceCol).value;
            console.log(`  Row ${r}: ${provinceVal}`);
        }

        // Count unique province values and count "Hồ Chí Minh"
        const provinceCount = {};
        let hcmCount = 0;
        let totalRows = 0;

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber > HEADER_ROW) {
                let provinceVal = row.getCell(provinceCol).value;

                // Handle different cell value types
                if (provinceVal && typeof provinceVal === 'object') {
                    if (provinceVal.richText) {
                        provinceVal = provinceVal.richText.map(rt => rt.text).join('');
                    } else if (provinceVal.text) {
                        provinceVal = provinceVal.text;
                    } else if (provinceVal.result) {
                        provinceVal = provinceVal.result;
                    }
                }

                const provinceStr = provinceVal ? provinceVal.toString().trim() : '(empty)';

                if (provinceStr) {
                    totalRows++;
                    provinceCount[provinceStr] = (provinceCount[provinceStr] || 0) + 1;

                    // Check for "Hồ Chí Minh" variants
                    const upper = provinceStr.toUpperCase();
                    if (upper.includes('HỒ CHÍ MINH') ||
                        upper.includes('HO CHI MINH') ||
                        upper.includes('HCM') ||
                        upper.includes('TP.HCM') ||
                        upper.includes('TP. HCM')) {
                        hcmCount++;
                    }
                }
            }
        });

        console.log(`\n--- Province Statistics ---`);
        console.log(`Total data rows: ${totalRows}`);
        console.log(`\nUnique provinces found:`);

        // Sort by count descending
        const sortedProvinces = Object.entries(provinceCount)
            .sort((a, b) => b[1] - a[1]);

        sortedProvinces.forEach(([province, count]) => {
            console.log(`  ${province}: ${count} rows`);
        });

        console.log(`\n--- Hồ Chí Minh Matching ---`);
        console.log(`Rows matching HCM variants: ${hcmCount}`);

    } catch (err) {
        console.error('Error:', err);
    }
}

inspectProvinceColumn();
