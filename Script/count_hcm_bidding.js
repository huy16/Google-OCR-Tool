const ExcelJS = require('exceljs');
const path = require('path');

const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const SHEET_NAME = 'General Data';
const HEADER_ROW = 2;

async function countHCMBidding() {
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

        // Count with BOTH filters
        let hcmOnlyCount = 0;
        let biddingOnlyCount = 0;
        let bothFiltersCount = 0;

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber > HEADER_ROW) {
                // Check Province (Col 2)
                let provinceVal = row.getCell(2).value;
                if (provinceVal && typeof provinceVal === 'object') {
                    provinceVal = provinceVal.text || provinceVal.result || '';
                }
                const provinceStr = provinceVal ? provinceVal.toString().trim().toUpperCase() : '';
                const isHCM = provinceStr.includes('HỒ CHÍ MINH');

                // Check Project Code (Col 52)
                const projectVal = row.getCell(52).value;
                const projectStr = projectVal ? projectVal.toString().toLowerCase() : '';
                const isBidding = projectStr.includes('2026') && projectStr.includes('bidding');

                if (isHCM) hcmOnlyCount++;
                if (isBidding) biddingOnlyCount++;
                if (isHCM && isBidding) bothFiltersCount++;
            }
        });

        console.log(`\n=== KẾT QUẢ ĐẾM ===`);
        console.log(`Chỉ lọc "Hồ Chí Minh":           ${hcmOnlyCount} dòng`);
        console.log(`Chỉ lọc "2026_bidding":          ${biddingOnlyCount} dòng`);
        console.log(`Lọc CẢ HAI (HCM + 2026_bidding): ${bothFiltersCount} dòng`);

    } catch (err) {
        console.error('Error:', err);
    }
}

countHCMBidding();
