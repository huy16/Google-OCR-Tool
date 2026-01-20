const puppeteer = require('../Script/node_modules/puppeteer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const EventEmitter = require('events');

class GeocoderService extends EventEmitter {
    constructor() {
        super();
        this.browser = null;
        this.isStopped = false;
    }

    async run(inputFilePath, outputDir, jobId, options = {}) {
        const {
            projectCodeFilter = '2026_bidding',
            provinceFilter = 'Hồ Chí Minh',
            districtFilter = '',
            surveyInfoFilter = ''
        } = options;

        this.emit('log', `🚀 Bắt đầu Job ${jobId}`);
        this.emit('log', `⚙️ Filter: Dự án="${projectCodeFilter}", Tỉnh="${provinceFilter}", Quận="${districtFilter}", Survey="${surveyInfoFilter}"`);
        this.emit('log', `📂 Input: ${path.basename(inputFilePath)}`);

        const OUTPUT_FILE_PATH = path.join(outputDir, `Result_${jobId}_v3.xlsx`);
        const CSV_BACKUP_PATH = path.join(outputDir, `Backup_${jobId}.csv`);
        const FINAL_EXCEL_PATH = path.join(outputDir, `Final_Result_${jobId}.xlsx`);

        // Ensure output dir exists
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // 1. Load Input Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inputFilePath);
        const sheet = workbook.getWorksheet('General Data') || workbook.worksheets[0];

        if (!sheet) {
            this.emit('error', 'Không tìm thấy sheet dữ liệu!');
            return;
        }

        // 2. Setup CSV Backup Header
        if (!fs.existsSync(CSV_BACKUP_PATH)) {
            fs.writeFileSync(CSV_BACKUP_PATH, 'STT Goc,Ma Kho 2,Tinh,Ten Cua Hang,Dia Chi,Link Map,Toa Do\n', 'utf8');
        }

        // 3. Map Columns
        let addressCol = 5, coordsCol = 10, linkCol = 11; // Defaults

        // Scan header
        const headerRow = sheet.getRow(2);
        headerRow.eachCell((cell, colNum) => {
            const val = cell.value ? cell.value.toString().trim() : '';
            if (val.includes('SIÊU THỊ/VĂN PHÒNG/KHO')) addressCol = colNum;
            if (val.includes('TỌA ĐỘ')) coordsCol = colNum;
            if (val.includes('Googgle Map')) linkCol = colNum;
        });

        this.emit('log', `✅ Cột dữ liệu: Address=${addressCol}, Coords=${coordsCol}, Link=${linkCol}`);

        // 4. Launch Browser
        this.emit('log', '🌍 Đang khởi động Google Chrome...');
        this.browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: ['--start-maximized']
        });
        const page = await this.browser.newPage();

        // 5. Processing Loop
        let processedCount = 0;
        const START_ROW = 3;

        // Count actual data rows
        let totalRowsToProcess = 0;
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber >= START_ROW) {
                // Optional: Check if row actually has data (e.g., STT exists)
                if (row.getCell(1).value) totalRowsToProcess++;
            }
        });
        this.emit('log', `📊 Tổng số dòng cần xử lý: ${totalRowsToProcess}`);

        for (let i = START_ROW; i <= sheet.rowCount; i++) {
            if (this.isStopped) break;

            const row = sheet.getRow(i);

            // --- FILTER LOGIC ---
            // 1. Project Code (Col 52)
            const projectCodeVal = row.getCell(52).value;
            const projectCodeStr = projectCodeVal ? projectCodeVal.toString().toLowerCase() : '';
            if (projectCodeFilter && !projectCodeStr.includes(projectCodeFilter.toLowerCase())) continue;

            // 2. Province (Col 2)
            const provinceVal = row.getCell(2).value;
            const provinceStr = (provinceVal && typeof provinceVal === 'object') ? (provinceVal.text || provinceVal.result || '') : (provinceVal ? provinceVal.toString().trim() : '');
            if (provinceFilter && !provinceStr.toUpperCase().includes(provinceFilter.toUpperCase())) continue;

            // 3. District (Col 3)
            const districtVal = row.getCell(3).value;
            const districtStr = (districtVal && typeof districtVal === 'object') ? (districtVal.text || districtVal.result || '') : (districtVal ? districtVal.toString().trim() : '');
            if (districtFilter && !districtStr.toUpperCase().includes(districtFilter.toUpperCase())) continue;

            // 4. Survey Info (Col 29)
            const surveyVal = row.getCell(29).value;
            const surveyStr = (surveyVal && typeof surveyVal === 'object') ? (surveyVal.text || surveyVal.result || '') : (surveyVal ? surveyVal.toString().trim() : '');

            if (surveyInfoFilter === '(Empty)') {
                if (surveyStr !== '') continue;
            } else if (surveyInfoFilter && !surveyStr.toUpperCase().includes(surveyInfoFilter.toUpperCase())) continue;

            // Valid row check
            const sttVal = row.getCell(1).value;
            if (!sttVal) continue;

            // --- END FILTER ---

            // Get Address
            let addressVal = row.getCell(addressCol).value;
            let address = '';
            if (addressVal && typeof addressVal === 'object') {
                if (addressVal.richText) address = addressVal.richText.map(rt => rt.text).join('');
                else if (addressVal.text) address = addressVal.text;
                else if (addressVal.hyperlink) address = addressVal.text || addressVal.hyperlink;
                else address = JSON.stringify(addressVal);
            } else {
                address = addressVal ? addressVal.toString() : '';
            }

            // Get Specific Address
            const specificAddressVal = row.getCell(12).value;
            let specificAddress = (specificAddressVal && typeof specificAddressVal === 'object')
                ? (specificAddressVal.text || specificAddressVal.result || '')
                : (specificAddressVal ? specificAddressVal.toString() : '');

            if (!specificAddress) specificAddress = address;

            // Build Query
            const searchQuery = `Bách Hóa Xanh ${specificAddress}`.trim();
            this.emit('log', `🔍 [Row ${i}] Tìm: ${searchQuery}`);

            // Calculate percentage based on actual total
            const percent = totalRowsToProcess > 0 ? Math.round(((processedCount + 1) / totalRowsToProcess) * 100) : 0;
            this.emit('progress', {
                row: i,
                percent: percent,
                message: searchQuery,
                current: processedCount + 1,
                total: totalRowsToProcess
            });

            // --- Puppeteer Logic ---
            try {
                await page.goto('https://www.google.com/maps', { waitUntil: 'networkidle2' });
                await this.delay(1000);

                await page.evaluate((query) => {
                    const box = document.querySelector('#searchboxinput') || document.querySelector('input[name="q"]');
                    if (box) { box.value = query; box.dispatchEvent(new Event('input', { bubbles: true })); }
                }, searchQuery);

                await page.keyboard.press('Enter');
                await this.delay(1500);

                // Smart Matching
                const shareBtnSelector = 'button[data-value="Share"]';
                let foundShare = false;

                try {
                    await page.waitForSelector(shareBtnSelector, { timeout: 4000 });
                    foundShare = true;
                } catch (e) {
                    // Smart Match Logic
                    const allResults = await page.$$('a[href^="https://www.google.com/maps/place"]');
                    if (allResults.length > 0) {
                        this.emit('log', `   Found ${allResults.length} results. Matching...`);

                        // Combined Expected Text
                        const districtVal = row.getCell(3).value;
                        const shopNameVal = row.getCell(5).value;
                        const combinedExpected = this.normalizeText(`${shopNameVal} ${districtVal} ${specificAddress}`);
                        const keywords = combinedExpected.split(' ').filter(w => w.length > 2);

                        let bestMatch = { element: null, score: 0 };

                        for (let idx = 0; idx < Math.min(allResults.length, 5); idx++) {
                            const resInfo = await page.evaluate(el => {
                                const p = el.closest('[jsaction]') || el.parentElement?.parentElement;
                                return p ? p.innerText : el.innerText;
                            }, allResults[idx]);

                            const normRes = this.normalizeText(resInfo);
                            let score = 0;
                            keywords.forEach(k => { if (normRes.includes(k)) score += k.length; });
                            if (normRes.includes('bach hoa xanh') || normRes.includes('bhx')) score += 20;

                            if (score > bestMatch.score) bestMatch = { element: allResults[idx], score };
                        }

                        if (bestMatch.element) {
                            await bestMatch.element.click();
                            await page.waitForSelector(shareBtnSelector, { timeout: 4000 });
                            foundShare = true;
                        }
                    }
                }

                // Get Link
                let finalLinkVal = 'KHÔNG TÌM THẤY';
                let coordsVal = '';

                if (foundShare) {
                    await page.click(shareBtnSelector);
                    await this.delay(500);

                    try {
                        await this.delay(800);
                        const link = await page.evaluate(() => {
                            const input = document.querySelector('input.vrsrZe') || document.querySelector('input[readonly]');
                            return input ? input.value : null;
                        });
                        if (link) finalLinkVal = link;
                    } catch (e) { }

                    await page.keyboard.press('Escape');
                    await this.delay(500);

                    // Coords extracted from URL (!3d!4d)
                    const url = page.url();
                    const match = url.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
                    if (match) coordsVal = `${match[1]}, ${match[2]}`;
                }

                this.emit('log', `   👉 Result: ${finalLinkVal}`);

                // --- Save to CSV ---
                const maKho2 = row.getCell(4).value ? row.getCell(4).value.toString() : '';
                const csvLine = [
                    i,
                    `"${maKho2}"`,
                    `"${provinceStr}"`,
                    `"${address.replace(/"/g, '""')}"`,
                    `"${specificAddress.replace(/"/g, '""')}"`,
                    `"${finalLinkVal}"`,
                    `"${coordsVal}"`
                ].join(',') + '\n';

                fs.appendFileSync(CSV_BACKUP_PATH, csvLine, 'utf8');
                processedCount++;

            } catch (err) {
                this.emit('log', `❌ Error Row ${i}: ${err.message}`);
            }
        } // End Loop

        await this.browser.close();

        // Convert CSV to Excel Final
        await this.convertCsvToExcel(CSV_BACKUP_PATH, FINAL_EXCEL_PATH);

        this.emit('complete', {
            processed: processedCount,
            file: FINAL_EXCEL_PATH
        });
    }

    normalizeText(text) {
        return (text || '').toLowerCase()
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
            .replace(/[^a-z0-9\s]/g, ' ')
            .replace(/\s+/g, ' ').trim();
    }

    delay(ms) {
        return new Promise(r => setTimeout(r, ms));
    }

    async convertCsvToExcel(csvPath, excelPath) {
        this.emit('log', '📦 Đang đóng gói file Excel cuối cùng...');
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Result');
        ws.columns = [
            { header: 'STT', key: 'stt', width: 10 },
            { header: 'MÃ KHO 2', key: 'mk', width: 15 },
            { header: 'Tỉnh', key: 't', width: 15 },
            { header: 'Tên Shop', key: 'n', width: 30 },
            { header: 'Địa Chỉ', key: 'a', width: 50 },
            { header: 'Link Map', key: 'l', width: 40 },
            { header: 'Tọa Độ', key: 'c', width: 20 },
        ];

        const content = fs.readFileSync(csvPath, 'utf8');
        const lines = content.split('\n');

        lines.forEach((line, idx) => {
            if (idx === 0 || !line.trim()) return;
            let cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
            cols = cols.map(c => c ? c.trim().replace(/^"|"$/g, '').replace(/""/g, '"') : '');
            if (cols.length >= 6) {
                ws.addRow({ stt: cols[0], mk: cols[1], t: cols[2], n: cols[3], a: cols[4], l: cols[5], c: cols[6] });
            }
        });

        await wb.xlsx.writeFile(excelPath);
        this.emit('log', '✅ Hoàn tất!');
    }

    async scanFile(filePath) {
        // Revert to Safe Mode (readFile) - Reliable for .xlsm
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        // Simple sheet logic (Version 1 style)
        let sheet = workbook.getWorksheet('General Data') || workbook.worksheets[0];
        if (!sheet) {
            throw new Error('No worksheets found in file');
        }

        const data = {
            projects: new Set(),
            provinces: new Set(),
            districts: new Set(),
            surveys: new Set()
        };

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber < 3) return;

            const extract = (cellIdx) => {
                const val = row.getCell(cellIdx).value;
                // Handle objects (rich text, formulas) or primitives
                let raw = (val && typeof val === 'object') ? (val.text || val.result || '') : val;
                // Force string and trim
                return raw ? String(raw).trim() : '';
            };

            const proj = extract(52);
            if (proj) data.projects.add(proj);

            const prov = extract(2);
            if (prov) data.provinces.add(prov);

            const dist = extract(3);
            if (dist) data.districts.add(dist);

            let surv = extract(29);
            if (!surv || surv.trim() === '') surv = '(Empty)';
            data.surveys.add(surv);
        });

        return {
            projects: [...data.projects].sort(),
            provinces: [...data.provinces].sort(),
            districts: [...data.districts].sort(),
            surveys: [...data.surveys].sort()
        };
    }
}

module.exports = new GeocoderService();
