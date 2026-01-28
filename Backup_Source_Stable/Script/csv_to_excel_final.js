const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const CSV_BACKUP_PATH = path.join(__dirname, '../1. Database/backup_results.csv');
const FINAL_EXCEL_PATH = path.join(__dirname, '../1. Database/Geocode_Results_Final.xlsx');

async function convertCsvToExcel() {
    console.log('Validating CSV and creating final Excel (Manual Parse)...');

    if (!fs.existsSync(CSV_BACKUP_PATH)) {
        console.error('Backup CSV not found!');
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Geocode Results');

    // Define Columns
    sheet.columns = [
        { header: 'STT Gốc', key: 'stt', width: 10 },
        { header: 'MÃ KHO 2', key: 'maKho', width: 15 },
        { header: 'Tỉnh', key: 'tinh', width: 15 },
        { header: 'Tên Cửa Hàng', key: 'ten', width: 30 },
        { header: 'Địa Chỉ Cụ Thể', key: 'diachi', width: 50 },
        { header: 'Google Map Link', key: 'link', width: 40 },
        { header: 'Tọa Độ', key: 'toado', width: 25 }
    ];

    try {
        const fileContent = fs.readFileSync(CSV_BACKUP_PATH, 'utf8');
        const lines = fileContent.split('\n');

        console.log(`Read ${lines.length} lines from CSV.`);

        let successCount = 0;
        lines.forEach((line, index) => {
            if (index === 0 || !line.trim()) return; // Skip Header & Empty lines

            // Basic CSV Splitter (Regex to handle commas inside quotes)
            let cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);

            // Cleanup quotes
            cols = cols.map(c => c ? c.trim().replace(/^"|"$/g, '').replace(/""/g, '"') : '');

            if (cols.length >= 6) {
                // CSV Header Structure:
                // 0: STT Goc
                // 1: Ma Kho 2
                // 2: Tinh
                // 3: Ten Cua Hang
                // 4: Dia Chi
                // 5: Link Map
                // 6: Toa Do
                sheet.addRow({
                    stt: cols[0],
                    maKho: cols[1],
                    tinh: cols[2],
                    ten: cols[3],
                    diachi: cols[4],
                    link: cols[5],
                    toado: cols[6] || ''
                });
                successCount++;
            }
        });

        // Format Header
        sheet.getRow(1).font = { bold: true };

        await workbook.xlsx.writeFile(FINAL_EXCEL_PATH);
        console.log(`✅ Successfully created: ${FINAL_EXCEL_PATH}`);
        console.log(`   Imported ${successCount} rows.`);

    } catch (err) {
        console.error('Error converting:', err);
    }
}

convertCsvToExcel();
