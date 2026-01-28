const geocoderService = require('../Web_App/geocoder_service');
const path = require('path');

async function test() {
    const filePath = "D:\\TOOL GOOGLE ANTIGRAVITY\\7. Tool OCR Google Map\\1. Database\\Test (Just View) CAS-MWG_Data.xlsm";
    console.log('Testing scan on:', filePath);
    try {
        const result = await geocoderService.scanFile(filePath);
        console.log('--- Result ---');
        console.log('Total Rows:', result.totalRows);
        if (result.filterData) {
            console.log('Filter Data Length:', result.filterData.length);
            console.log('Sample Item:', JSON.stringify(result.filterData[0], null, 2));
        } else {
            console.log('ERROR: filterData is missing!');
            console.log('Keys returned:', Object.keys(result));
        }
    } catch (e) {
        console.error('Error:', e);
    }
}

test();
