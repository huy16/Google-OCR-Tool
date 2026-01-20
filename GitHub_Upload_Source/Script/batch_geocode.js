const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Helper function for delay (since page.waitForTimeout is deprecated)
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Configuration
const EXCEL_FILE_PATH = path.join(__dirname, '../1. Database/(Just View) CAS-MWG_Data.xlsm');
const OUTPUT_FILE_PATH = path.join(__dirname, '../1. Database/Geocode_Results_Simple_v3.xlsx'); // NEW: Version 3 to avoid file lock v2
const CSV_BACKUP_PATH = path.join(__dirname, '../1. Database/backup_results.csv'); // BACKUP CSV
const SHEET_NAME = 'General Data';
const HEADER_ROW_INDEX = 2;
const START_ROW = 3;

// Column Headers to map (will search for these)
const COL_ADDRESS = 'ĐỊA CHỈ CỤ THỂ';
const COL_COORDS = 'Tọa độ';
const COL_MAP_LINK = 'Googgle Map';

async function main() {
    console.log('Starting Batch Geocoder (Safe Mode)...');

    // 1. Load Input Excel (Read Only)
    if (!fs.existsSync(EXCEL_FILE_PATH)) {
        console.error(`File not found: ${EXCEL_FILE_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        console.log('Input Workbook loaded.');
    } catch (err) {
        console.error('Error reading Excel file:', err);
        return;
    }

    // Prepare Output Workbook (New Clean File)
    const outWorkbook = new ExcelJS.Workbook();
    let outSheet;

    // Check if CSV exists, if not create header
    if (!fs.existsSync(CSV_BACKUP_PATH)) {
        fs.writeFileSync(CSV_BACKUP_PATH, 'STT Goc,Ma Kho 2,Tinh,Ten Cua Hang,Dia Chi,Link Map,Toa Do\n', 'utf8');
        console.log('Created new CSV backup file.');
    }

    // Check if output file exists to append/resume
    let resumeRowMap = new Set();
    if (fs.existsSync(OUTPUT_FILE_PATH)) {
        console.log('Found existing output file. Loading to resume...');
        await outWorkbook.xlsx.readFile(OUTPUT_FILE_PATH);
        outSheet = outWorkbook.getWorksheet('Results');
        if (outSheet) {
            // Load processed rows into Set
            outSheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Skip header
                    const originalRowIndex = row.getCell(1).value;
                    if (originalRowIndex) resumeRowMap.add(parseInt(originalRowIndex));
                }
            });
            console.log(`Resuming: Found ${resumeRowMap.size} already processed rows.`);
        }
    }

    if (!outSheet) {
        outSheet = outWorkbook.addWorksheet('Results');
        outSheet.columns = [
            { header: 'STT Gốc', key: 'rowIndex', width: 12 },
            { header: 'MÃ KHO 2', key: 'maKho2', width: 15 },
            { header: 'Tỉnh', key: 'province', width: 18 },
            { header: 'Mã Dự Án', key: 'pCode', width: 20 },
            { header: 'Tên Cửa Hàng', key: 'shop', width: 35 },
            { header: 'Địa Chỉ Cụ Thể', key: 'addr', width: 50 },
            { header: 'Google Map Link', key: 'link', width: 45 },
            { header: 'Tọa Độ', key: 'coords', width: 25 }
        ];
    }

    // 2. Find Sheet
    const sheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];
    if (!sheet) {
        console.error('Sheet not found.');
        return;
    }
    console.log(`Using sheet: ${sheet.name}`);

    // 3. Map Columns
    // Let's scan the header row (e.g. row 2)
    let addressCol = null;
    let coordsCol = null;
    let linkCol = null;

    const headerRow = sheet.getRow(HEADER_ROW_INDEX);
    headerRow.eachCell((cell, colNumber) => {
        const val = cell.value ? cell.value.toString().trim() : '';
        if (val.includes('SIÊU THỊ/VĂN PHÒNG/KHO')) addressCol = colNumber; // Col 5 per inspection
        if (val.includes('TỌA ĐỘ')) coordsCol = colNumber;             // Col 10 per inspection
        if (val.includes('Googgle Map')) linkCol = colNumber;          // Col 11 per inspection (Typo in header "Googgle")
    });

    // Fallback/Hardcode if auto-detect fails (Safe for General Data)
    if (!addressCol) addressCol = 5;
    if (!coordsCol) coordsCol = 10;
    if (!linkCol) linkCol = 11;

    if (!addressCol || !coordsCol || !linkCol) {
        console.error('Do not find columns:', { addressCol, coordsCol, linkCol });
        return;
    }

    console.log(`Columns mapped: Address=${addressCol}, Coords=${coordsCol}, Link=${linkCol}`);

    // 4. Launch Browser
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
        args: ['--start-maximized']
    });
    const page = await browser.newPage();

    // 5. Iterate
    let processedCount = 0;
    const maxItemsToProcess = 500; // Xử lý hết 491 cửa hàng HCM
    const FETCH_COORDS = false; // Keep coordinate fetching paused

    // Iterate through all rows, but stop when we have processed enough items
    for (let i = START_ROW; i <= sheet.rowCount; i++) {
        if (processedCount >= maxItemsToProcess) {
            console.log(`\nReached limit of ${maxItemsToProcess} items. Stopping test run.`);
            break;
        }
        const row = sheet.getRow(i);

        // Filter by Project Code "2026_bidding" (Col 52)
        const projectCodeVal = row.getCell(52).value; // MÃ DỰ ÁN is Col 52
        const projectCodeStr = projectCodeVal ? projectCodeVal.toString().toLowerCase() : '';

        if (!projectCodeStr.includes('2026') || !projectCodeStr.includes('bidding')) {
            continue;
        }

        // Filter by Province "Hồ Chí Minh" (Col 2 = TỈNH)
        const provinceVal = row.getCell(2).value;
        let provinceStr = '';
        if (provinceVal && typeof provinceVal === 'object') {
            provinceStr = provinceVal.text || provinceVal.result || '';
        } else {
            provinceStr = provinceVal ? provinceVal.toString().trim() : '';
        }

        // Chỉ xử lý Hồ Chí Minh
        if (!provinceStr.toUpperCase().includes('HỒ CHÍ MINH')) {
            continue;
        }

        // Check if already processed (Resume)
        if (resumeRowMap.has(i)) {
            console.log(`Skipping Row ${i} (Already processed)`);
            continue;
        }

        // STT Check (Col 1)
        const sttVal = row.getCell(1).value;
        if (sttVal === null || sttVal === undefined || sttVal === '') {
            continue;
        }

        // Handle Rich Text / Hyperlink objects in ExcelJS
        let addressVal = row.getCell(addressCol).value;
        let address = '';
        console.log(`Row ${i} Val:`, JSON.stringify(addressVal)); // Debug log

        if (addressVal && typeof addressVal === 'object') {
            if (addressVal.richText && Array.isArray(addressVal.richText)) {
                address = addressVal.richText.map(rt => rt.text).join('');
            } else if (addressVal.text) {
                address = addressVal.text;
            } else if (addressVal.result) {
                // Check if result is error object
                if (typeof addressVal.result === 'object' && addressVal.result.error) {
                    console.log(`   -> Skipped Row ${i} (Formula Error: ${addressVal.result.error})`);
                    continue;
                }
                address = addressVal.result;
            } else if (addressVal.hyperlink) {
                address = addressVal.text || addressVal.hyperlink;
            } else {
                address = JSON.stringify(addressVal);
            }
        } else {
            address = addressVal ? addressVal.toString() : '';
        }

        // Skip empty addresses
        if (!address) continue;

        // Address Strategy: Use Column L (Specific Address) + "Bách Hóa Xanh"
        // Col L often has full address: "Đường Nguyễn Văn Linh, Phường Tân Châu..."
        const specificAddressVal = row.getCell(12).value; // Col 12 = L
        let specificAddress = '';

        if (specificAddressVal && typeof specificAddressVal === 'object') {
            specificAddress = specificAddressVal.text || specificAddressVal.result || '';
        } else {
            specificAddress = specificAddressVal ? specificAddressVal.toString() : '';
        }

        // Fallback to Col E if Col L is empty, but clean it
        if (!specificAddress) {
            specificAddress = address; // Fallback to 'address' var which came from Col 5
        }

        // Construct Query: "Bách Hóa Xanh" + Specific Address
        // This is more accurate than combining E + L because E often has internal codes.
        searchQuery = `Bách Hóa Xanh ${specificAddress}`.trim();

        // Check if already has data (optional: skip if exists?)
        const existingLink = row.getCell(linkCol).value;
        if (existingLink && existingLink.toString().includes('google')) {
            // console.log(`Row ${i}: Already has link. Skipping.`);
            // continue;
        }

        console.log(`Processing Row ${i}: ${searchQuery}`);

        try {
            await page.goto('https://www.google.com/maps', { waitUntil: 'networkidle2' });

            // Wait for page to fully load
            await delay(1000); // Giảm từ 2000

            // Search using JavaScript (more reliable than typing Vietnamese characters)
            await page.evaluate((query) => {
                const searchBox = document.querySelector('#searchboxinput') || document.querySelector('input[name="q"]');
                if (searchBox) {
                    searchBox.value = query;
                    searchBox.dispatchEvent(new Event('input', { bubbles: true }));
                    searchBox.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }, searchQuery);

            // Click search button or press Enter
            await page.keyboard.press('Enter');

            // Wait for results
            await delay(1500); // Giảm từ 3000

            // Wait for the "Share" button. It usually appears in the side panel for a specific place.
            // Selector for share button: usually has text "Chia sẻ" or data-value="Share"
            // We might need to select the first result if multiple appear.

            // Check if we are on a "search results list" page or a "place details" page.
            // If "place details", we see actions like Directions, Save, Share.
            // Share button selector: button[data-value="Share"] or similar.

            const shareBtnSelector = 'button[data-value="Share"]';
            let foundShare = false;

            try {
                await page.waitForSelector(shareBtnSelector, { timeout: 5000 });
                foundShare = true;
            } catch (e) {
                // === SMART ADDRESS MATCHING ===
                // Lấy tất cả kết quả tìm kiếm và so khớp với địa chỉ gốc
                const allResults = await page.$$('a[href^="https://www.google.com/maps/place"]');

                if (allResults.length > 0) {
                    console.log(`   -> Tìm thấy ${allResults.length} kết quả, đang so khớp...`);

                    // Chuẩn bị địa chỉ gốc để so sánh (lowercase, bỏ dấu)
                    const normalizeText = (text) => {
                        return text.toLowerCase()
                            .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // Bỏ dấu
                            .replace(/[^a-z0-9\s]/g, ' ') // Chỉ giữ chữ và số
                            .replace(/\s+/g, ' ').trim();
                    };

                    // Lấy thêm cột C (Quận/Huyện) và cột E (Tên cửa hàng) để so khớp
                    const districtVal = row.getCell(3).value; // Col C = QUẬN/HUYỆN
                    const shopNameVal = row.getCell(5).value; // Col E = SIÊU THỊ/VĂN PHÒNG/KHO

                    const districtStr = districtVal ? districtVal.toString() : '';
                    const shopNameStr = shopNameVal ? shopNameVal.toString() : '';

                    // Kết hợp tất cả thông tin để so khớp
                    const combinedExpected = normalizeText(`${shopNameStr} ${districtStr} ${specificAddress}`);
                    const expectedKeywords = combinedExpected.split(' ').filter(w => w.length > 2);

                    console.log(`   -> Từ khóa so khớp: ${expectedKeywords.slice(0, 8).join(', ')}...`);

                    let bestMatch = { element: null, score: 0, name: '' };

                    for (let idx = 0; idx < Math.min(allResults.length, 5); idx++) {
                        const result = allResults[idx];

                        // Lấy text của kết quả (tên địa điểm + địa chỉ)
                        const resultText = await page.evaluate(el => {
                            // Tìm container cha chứa tên và địa chỉ
                            const parent = el.closest('[jsaction]') || el.parentElement?.parentElement;
                            return parent ? parent.innerText : el.innerText;
                        }, result);

                        const normalizedResult = normalizeText(resultText);

                        // Tính điểm so khớp
                        let score = 0;
                        for (const keyword of expectedKeywords) {
                            if (normalizedResult.includes(keyword)) {
                                score += keyword.length; // Từ dài hơn = điểm cao hơn
                            }
                        }

                        // Ưu tiên kết quả có "Bach Hoa Xanh" hoặc "BHX"
                        if (normalizedResult.includes('bach hoa xanh') || normalizedResult.includes('bhx')) {
                            score += 20;
                        }

                        console.log(`      Kết quả ${idx + 1}: Score=${score} - ${resultText.substring(0, 50)}...`);

                        if (score > bestMatch.score) {
                            bestMatch = { element: result, score: score, name: resultText.substring(0, 50) };
                        }
                    }

                    // Click kết quả khớp nhất (hoặc cái đầu tiên nếu không có kết quả nào khớp)
                    const selectedResult = bestMatch.element || allResults[0];
                    console.log(`   -> Chọn kết quả khớp nhất: Score=${bestMatch.score}`);

                    await selectedResult.click();
                    await page.waitForSelector(shareBtnSelector, { timeout: 5000 });
                    foundShare = true;
                }
            }

            if (foundShare) {
                let lat = '', long = '';
                let linkVal = '';

                // Step 1: Click Share button and get the link FIRST
                await page.click(shareBtnSelector);
                await delay(500); // Giảm từ 1000

                // Wait for modal and get link
                try {
                    await delay(800); // Giảm từ 1500

                    // Try multiple selectors for the share link
                    linkVal = await page.evaluate(() => {
                        // Method 1: Look for input with the link
                        const input = document.querySelector('input.vrsrZe') ||
                            document.querySelector('input[readonly]') ||
                            document.querySelector('input[value*="maps.app.goo.gl"]');
                        if (input && input.value) return input.value;

                        // Method 2: Find any element containing the maps link text
                        const allElements = document.querySelectorAll('*');
                        for (const el of allElements) {
                            const text = el.innerText || el.textContent || '';
                            if (text.includes('maps.app.goo.gl') && text.startsWith('https://')) {
                                // Make sure it's just the link, not a longer text
                                const match = text.match(/https:\/\/maps\.app\.goo\.gl\/[a-zA-Z0-9]+/);
                                if (match) return match[0];
                            }
                        }
                        return null;
                    });

                    if (!linkVal) {
                        console.log(`   -> Could not find share link in modal`);
                    }
                } catch (e) {
                    console.log(`   -> Could not get share link: ${e.message}`);
                }

                // Close Share modal
                await page.keyboard.press('Escape');
                await delay(500);

                // Step 2: ZOOM INTO the pin and right-click to get accurate coordinates
                if (FETCH_COORDS) {
                    try {
                        // After Share modal closes, the place should still be selected
                        // We need to zoom into the pin to center it on screen

                        // Method 1: Click on the place card/header to ensure focus on the place
                        const placeHeader = await page.$('h1.DUwDvf, h1.fontHeadlineLarge');
                        if (placeHeader) {
                            await placeHeader.click();
                            await delay(500);
                        }

                        // Method 2: Zoom into the map to center on the selected place
                        // Press "+" key multiple times to zoom in (this centers on selected place)
                        const canvas = await page.$('canvas');
                        if (canvas) {
                            // Click on canvas first to focus it
                            const box = await canvas.boundingBox();
                            if (box) {
                                await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
                                await delay(300);

                                // Zoom in by pressing + key 3 times to get closer to the pin
                                for (let z = 0; z < 3; z++) {
                                    await page.keyboard.press('+');
                                    await delay(300);
                                }
                                await delay(1000); // Wait for zoom animation

                                // Now the selected place's pin should be centered
                                // Right-click at the center of the map (where pin is)
                                await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' });
                                await delay(1500);

                                // Get coordinates from context menu using discovered selector
                                const coordsText = await page.evaluate(() => {
                                    const firstMenuItem = document.querySelector('div.fxNQSd div.mLuXec');
                                    if (firstMenuItem) {
                                        const text = firstMenuItem.innerText || firstMenuItem.textContent || '';
                                        if (/^-?\d+\.\d+,\s*-?\d+\.\d+$/.test(text.trim())) {
                                            return text.trim();
                                        }
                                    }
                                    return null;
                                });

                                if (coordsText) {
                                    const parts = coordsText.split(',').map(s => s.trim());
                                    lat = parts[0];
                                    long = parts[1];
                                }

                                // Close context menu
                                await page.keyboard.press('Escape');
                                await delay(300);
                            }
                        }

                        // PRIMARY METHOD: Extract coordinates from URL !3d!4d pattern
                        // This pattern gives EXACT place pin coordinates (not view center)
                        // URL format: ...data=!3d10.7934252!4d106.5762635...
                        if (!lat || !long) {
                            const currentUrl = page.url();
                            // !3d[LAT]!4d[LONG] contains the EXACT place coordinates
                            const placeMatch = currentUrl.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
                            if (placeMatch) {
                                lat = placeMatch[1];
                                long = placeMatch[2];
                                console.log(`   -> Extracted from URL (!3d!4d): ${lat}, ${long}`);
                            } else {
                                // Fallback: @lat,long is view center, less accurate but usable
                                const viewMatch = currentUrl.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
                                if (viewMatch) {
                                    lat = viewMatch[1];
                                    long = viewMatch[2];
                                    console.log(`   -> Extracted from URL (@view): ${lat}, ${long}`);
                                }
                            }
                        }
                    } catch (coordErr) {
                        console.log(`   -> Could not get coords: ${coordErr.message}`);
                    }
                } // End if (FETCH_COORDS)

                console.log(`   -> Found: ${linkVal} (${lat}, ${long})`);

                // Validate Coordinates for HCM
                let note = "";
                if (lat && long) {
                    const latNum = parseFloat(lat);
                    const longNum = parseFloat(long);
                    const addressUpper = String(addressVal || "").toUpperCase();

                    // Basic check for HCM (Range approx: Lat 10.3-11.3, Long 106.3-107.0)
                    if (addressUpper.includes("HCM") || addressUpper.includes("HỒ CHÍ MINH") || addressUpper.includes("HO CHI MINH")) {
                        if (latNum < 10.3 || latNum > 11.3 || longNum < 106.3 || longNum > 107.0) {
                            console.warn(`   ⚠️ WARNING: Coordinates (${lat}, ${long}) seem outside HCM!`);
                            note = " [CHECK: Outside HCM]";
                        }
                    }
                }

                // Update Output File (Append Row)
                const finalLinkVal = (linkVal || "") + note;
                const coordsVal = (lat && long) ? `${lat}, ${long}` : '';

                // Get MÃ KHO 2 from Col 4
                const maKho2Val = row.getCell(4).value;
                const maKho2Str = maKho2Val ? maKho2Val.toString() : '';

                outSheet.addRow({
                    rowIndex: i,
                    maKho2: maKho2Str,
                    province: provinceStr,
                    pCode: projectCodeStr,
                    shop: address,
                    addr: specificAddress,
                    link: finalLinkVal,
                    coords: coordsVal
                });

                // === BACKUP TO CSV IMMIDIATELY ===
                try {
                    const csvLine = [
                        i, // STT Goc
                        `"${maKho2Str}"`,
                        `"${provinceStr}"`,
                        `"${address.replace(/"/g, '""')}"`, // Ten Cua Hang
                        `"${specificAddress.replace(/"/g, '""')}"`, // Dia Chi
                        `"${finalLinkVal}"`,
                        `"${coordsVal}"`
                    ].join(',') + '\n';

                    fs.appendFileSync(CSV_BACKUP_PATH, csvLine, 'utf8');
                } catch (csvErr) {
                    console.error('Error writing to CSV backup:', csvErr);
                }
                // =================================

                // Close modal (Esc)
                await page.keyboard.press('Escape');

                processedCount++;

                // Save periodically every 5 rows
                if (processedCount % 5 === 0) {
                    await outWorkbook.xlsx.writeFile(OUTPUT_FILE_PATH);
                    console.log('   [Saved Progress to Geocode_Results_Simple_v2.xlsx]');
                }

            } else {
                console.log('   -> Không tìm thấy địa điểm/nút Share.');

                // Get MÃ KHO 2 from Col 4
                const maKho2ValFail = row.getCell(4).value;
                const maKho2StrFail = maKho2ValFail ? maKho2ValFail.toString() : '';

                outSheet.addRow({
                    rowIndex: i,
                    maKho2: maKho2StrFail,
                    province: provinceStr,
                    pCode: projectCodeStr,
                    shop: address,
                    addr: specificAddress,
                    link: 'KHÔNG TÌM THẤY',
                    coords: ''
                });

                // === BACKUP TO CSV FAILS TOO ===
                try {
                    const csvLine = [
                        i, // STT Goc
                        `"${maKho2StrFail}"`,
                        `"${provinceStr}"`,
                        `"${address.replace(/"/g, '""')}"`,
                        `"${specificAddress.replace(/"/g, '""')}"`,
                        '"KHÔNG TÌM THẤY"',
                        '""'
                    ].join(',') + '\n';

                    fs.appendFileSync(CSV_BACKUP_PATH, csvLine, 'utf8');
                } catch (csvErr) {
                    console.error('Error writing to CSV backup:', csvErr);
                }
                // ==============================
            }

        } catch (err) {
            console.error(`   -> Error processing row ${i}:`, err.message);
        }

    }

    // Final Save
    await outWorkbook.xlsx.writeFile(OUTPUT_FILE_PATH);
    console.log(`Done! Saved to: ${OUTPUT_FILE_PATH}`);

    await browser.close();
}

main();
