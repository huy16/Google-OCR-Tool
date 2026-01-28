const puppeteer = require('puppeteer');
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

    stop() {
        this.isStopped = true;
        console.log('!!! SERVICE STOP METHOD CALLED. isStopped = true !!!');
        this.emit('log', '🛑 Đang dừng xử lý...');
    }

    async run(inputFilePath, outputDir, jobId, options = {}) {
        try {
            this.isStopped = false; // Reset flag
            const {
                projectCodeFilter = '2026_bidding',
                provinceFilter = 'Hồ Chí Minh',
                districtFilter = '',
                surveyInfoFilter = '',
                headless = true
            } = options;

            this.emit('log', `🚀 Bắt đầu Job ${jobId}`);
            // ... (rest of logging) ...
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
            this.emit('log', `🌍 Đang khởi động Google Chrome... (Headless: ${headless})`);

            const launchArgs = ['--no-sandbox', '--disable-setuid-sandbox'];
            let defViewport = { width: 800, height: 600 };

            if (!headless) {
                launchArgs.push('--start-maximized');
                defViewport = null; // Let browser determine size
            }

            this.browser = await puppeteer.launch({
                headless: headless,
                defaultViewport: defViewport,
                executablePath: this.getChromiumExecPath(), // Custom path resolution
                args: [...launchArgs, '--disable-gpu', '--disable-dev-shm-usage', '--no-first-run', '--mute-audio']
            });
            const page = await this.browser.newPage();

            // Optimize: Block images/fonts/css
            await page.setRequestInterception(true);
            page.on('request', (req) => {
                const type = req.resourceType();
                if (['image', 'stylesheet', 'font', 'media'].includes(type)) {
                    req.abort();
                } else {
                    req.continue();
                }
            });

            // Navigate to Google Maps
            this.emit('log', '🌍 Đang truy cập Google Maps...');
            await page.goto('https://www.google.com/maps?hl=vi', { waitUntil: 'networkidle2' });

            try {
                const searchBoxSel = '#searchboxinput';
                await page.waitForSelector(searchBoxSel, { timeout: 10000 });
            } catch (e) {
                this.emit('log', '⚠️ Warning: Map load slow or selector changed.');
            }
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
                    // Debug logs
                    this.emit('log', `   Wait content load...`);

                    // Native Puppeteer Interaction (More Robust for PKG)
                    const searchBoxSel = '#searchboxinput';
                    const fallbackSel = 'input[name="q"]'; // Common fallback

                    try {
                        // Click target
                        await page.waitForSelector(searchBoxSel, { timeout: 8000 }).catch(() => { });
                        const targetSel = (await page.$(searchBoxSel)) ? searchBoxSel : fallbackSel;

                        try { await page.click(targetSel); } catch (e) { }

                        // --- OPTIMIZED INPUT (Instant) ---
                        const inputFn = new Function('sel', 'query', `
                            const box = document.querySelector(sel);
                            if (box) { 
                                box.value = query; 
                                box.dispatchEvent(new Event('input', { bubbles: true })); 
                                box.dispatchEvent(new Event('change', { bubbles: true })); 
                                box.focus(); 
                            }
                        `);
                        await page.evaluate(inputFn, targetSel, searchQuery);
                        await this.delay(200);
                        await page.keyboard.press('Enter');
                        // ---------------------------------

                    } catch (errInput) {
                        this.emit('log', `⚠️ Input Error: ${errInput.message}`);
                    }

                    // OPTIMIZED WAIT: Wait for result or share button instead of hard wait
                    try {
                        const shareBtnSel = 'button[data-value="Share"]';
                        const resultLinkSel = 'a[href^="https://www.google.com/maps/place"]';

                        await page.waitForFunction(
                            (s1, s2) => document.querySelector(s1) || document.querySelector(s2),
                            { timeout: 3000 },
                            shareBtnSel,
                            resultLinkSel
                        );
                    } catch (e) { /* Check failed, proceed to standard checks */ }


                    // Smart Matching & Share Click
                    let foundShare = false;

                    // 1. Direct Share Button (Address loaded directly)
                    const directShareSelectors = [
                        'button[data-value="Share"]',
                        'button[aria-label="Share"]',
                        'button[aria-label="Chia sẻ"]',
                        '[data-tooltip="Chia sẻ"]',
                        '[data-tooltip="Share"]'
                    ];

                    for (const sel of directShareSelectors) {
                        const btn = await page.$(sel);
                        if (btn) {
                            try {
                                await btn.click();
                                foundShare = true;
                                this.emit('log', `   ✨ Found Share button directly: ${sel}`);
                                break;
                            } catch (e) { }
                        }
                    }

                    // 2. If not found, look for Result List logic
                    if (!foundShare) {
                        try {
                            // Wait briefly for Share button if it's loading
                            await page.waitForSelector('button[data-value="Share"]', { timeout: 2000 });
                            foundShare = true;
                            await page.click('button[data-value="Share"]');
                        } catch (e) { /* Ignore */ }
                    }

                    if (!foundShare) {
                        // Smart Match Logic (Fallback for list view)
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
                                // Use new Function string to survive pkg stripping
                                const extractFn = new Function('el', `
                                    const p = el.closest('[jsaction]') || el.parentElement?.parentElement;
                                    return p ? p.innerText : el.innerText;
                                `);
                                const resInfo = await page.evaluate(extractFn, allResults[idx]);

                                const normRes = this.normalizeText(resInfo);
                                let score = 0;
                                keywords.forEach(k => { if (normRes.includes(k)) score += k.length; });
                                if (normRes.includes('bach hoa xanh') || normRes.includes('bhx')) score += 20;

                                if (score > bestMatch.score) bestMatch = { element: allResults[idx], score };
                            }

                            if (bestMatch.element) {
                                await bestMatch.element.click();
                                // await this.delay(2000); // Wait for details panel
                                // Optimized: Wait for Share button to appear
                                try {
                                    await page.waitForSelector('button[data-value="Share"]', { timeout: 3000 });
                                } catch (e) { }

                                // Retry finding Share button after clicking result
                                for (const sel of directShareSelectors) {
                                    try {
                                        await page.waitForSelector(sel, { timeout: 1500 });
                                        await page.click(sel);
                                        foundShare = true;
                                        this.emit('log', `   ✨ Found Share button after click: ${sel}`);
                                        break;
                                    } catch (e) { }
                                }

                                if (!foundShare) this.emit('log', '⚠️ Could not find Share button after clicking result.');
                            }
                        }
                    }

                    // Get Link
                    let finalLinkVal = 'KHÔNG TÌM THẤY';
                    let coordsVal = '';

                    if (foundShare) {
                        // processing copy link...
                        await this.delay(1000); // Wait for modal animation

                        try {
                            // 1. Click "Copy Link" / "Sao chép" button
                            const copyBtnSelectors = [
                                'button[data-tooltip="Sao chép đường liên kết"]',
                                'button[aria-label="Sao chép đường liên kết"]',
                                'button[data-tooltip="Copy link"]',
                                'button[aria-label="Copy link"]',
                                '.yA7sBe button' // Generic class sometimes used in modal
                            ];

                            // Try strict selectors first
                            let clickedCopy = false;
                            for (const sel of copyBtnSelectors) {
                                const btn = await page.$(sel);
                                if (btn) {
                                    await btn.click();
                                    clickedCopy = true;
                                    break;
                                }
                            }

                            // If strict failed, try text based (simpler)
                            if (!clickedCopy) {
                                const clickCopyFx = new Function(`
                                    const btns = Array.from(document.querySelectorAll('button'));
                                    for (const b of btns) {
                                        const t = b.innerText.toLowerCase();
                                        if (t.includes('sao chép') || t.includes('copy link')) {
                                            b.click();
                                            return true;
                                        }
                                    }
                                    return false;
                                `);
                                await page.evaluate(clickCopyFx);
                            }

                            await this.delay(500);

                            // 2. Read the Link from the readonly input box (Reliable)
                            const linkFn = new Function(`
                                const input = document.querySelector('input.vrsrZe') || document.querySelector('input[readonly]');
                                return input ? input.value : null;
                            `);
                            const link = await page.evaluate(linkFn);
                            if (link) finalLinkVal = link;
                        } catch (e) {
                            this.emit('log', `⚠️ Share/Copy Error: ${e.message}`);
                        }

                    } // End if(foundShare)

                    // FALLBACK: Use Current URL if Link is still invalid
                    // FALLBACK REMOVED: Strict Mode
                    if (!finalLinkVal) finalLinkVal = 'KHÔNG TÌM THẤY';

                    await page.keyboard.press('Escape');
                    await this.delay(500);

                    // Coords extracted from URL (!3d!4d)
                    const url = page.url();
                    const match = url.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
                    if (match) coordsVal = `${match[1]}, ${match[2]}`;

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
        } catch (fatalError) {
            console.error('❌ FATAL ERROR IN RUN:', fatalError);
            this.emit('error', 'Lỗi nghiêm trọng: ' + fatalError.message);
            this.emit('log', '❌ Lỗi nghiêm trọng: ' + fatalError.message);
        }
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

    getChromiumExecPath() {
        // 1. Priority: Portable Chrome next to EXE
        const portablePath = path.join(process.cwd(), 'chrome-win', 'chrome.exe');
        const fs = require('fs');
        if (fs.existsSync(portablePath)) {
            console.log('Using portable Chrome:', portablePath);
            return portablePath;
        }

        // 2. Fallback: System Chrome (Windows)
        const systemPaths = [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
            process.env.LOCALAPPDATA + '\\Google\\Chrome\\Application\\chrome.exe'
        ];

        for (const p of systemPaths) {
            if (fs.existsSync(p)) {
                console.log('Using system Chrome:', p);
                return p;
            }
        }

        // 3. Fallback: Edge (if Chrome not found)
        const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
        if (fs.existsSync(edgePath)) {
            console.log('Using Microsoft Edge:', edgePath);
            return edgePath;
        }

        console.log('Chrome not found. Letting Puppeteer try default (might fail in EXE)...');
        return undefined; // Puppeteer default
    }

    async scanFile(filePath) {
        this.emit('log', `Scanning file: ${filePath}`);

        if (!fs.existsSync(filePath)) {
            throw new Error(`File not found: ${filePath}`);
        }

        // Revert to Safe Mode (readFile) - Reliable for .xlsm
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        // Simple sheet logic (Version 1 style)
        let sheet = workbook.getWorksheet('General Data') || workbook.worksheets[0];
        if (!sheet) {
            throw new Error('No worksheets found in file');
        }

        const uniqueCombinations = new Map();
        let totalRows = 0;

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber < 3) return;
            // Check if valid row (has STT or Project Code)
            if (!row.getCell(1).value && !row.getCell(52).value) return;

            totalRows++;

            const extract = (cellIdx) => {
                const val = row.getCell(cellIdx).value;
                let raw = (val && typeof val === 'object') ? (val.text || val.result || '') : val;
                return raw ? String(raw).trim() : '';
            };

            const proj = extract(52);
            const prov = extract(2);
            const dist = extract(3);
            let surv = extract(29);
            if (!surv || surv.trim() === '') surv = '(Empty)';

            // Create a unique key for this combination
            const key = JSON.stringify({ proj, prov, dist, surv });
            uniqueCombinations.set(key, (uniqueCombinations.get(key) || 0) + 1);
        });

        const filterData = [];
        uniqueCombinations.forEach((count, key) => {
            const obj = JSON.parse(key);
            obj.count = count;
            filterData.push(obj);
        });

        return {
            totalRows: totalRows,
            filterData: filterData
        };
    }
}

module.exports = new GeocoderService();
