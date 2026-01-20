const ExcelJS = require('exceljs');
const path = require('path');

(async () => {
    const filePath = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
    console.log(`Reading: ${filePath}`);

    const options = {
        sharedStrings: 'cache',
        hyperlinks: 'ignore',
        styles: 'ignore'
    };

    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, options);

    for await (const worksheetReader of workbookReader) {
        console.log(`Worksheet: ${worksheetReader.id}`);
        let count = 0;
        for await (const row of worksheetReader) {
            if (count > 5) break;
            console.log(`Row ${row.number}:`);
            // Try accessing Col 2 (Province) and Col 52 (Project)
            try {
                // Method 1: getCell
                const c2 = row.getCell(2).value;
                const c52 = row.getCell(52).value;
                console.log(`   [getCell] Col 2: ${JSON.stringify(c2)}, Col 52: ${JSON.stringify(c52)}`);
            } catch (e) { console.log('   [getCell] Error:', e.message); }

            try {
                // Method 2: values array
                // Note: values array is 1-based (index 0 is null/undefined)
                if (row.values && row.values.length > 52) {
                    console.log(`   [values] Col 2: ${JSON.stringify(row.values[2])}, Col 52: ${JSON.stringify(row.values[52])}`);
                } else {
                    console.log(`   [values] Length: ${row.values ? row.values.length : 'undefined'}`);
                }
            } catch (e) { console.log('   [values] Error:', e.message); }

            count++;
        }
        break; // Only first sheet
    }
})();
