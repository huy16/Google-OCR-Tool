const fs = require('fs');
const path = require('path');

async function testScan() {
    const formData = new FormData();
    // Use a real excel file if possible, or try a dummy one to see if it even reaches the server
    // Ideally we need an excel file to pass the file type check/parsing
    // But let's see if we can just ping it first

    // We'll create a simple buffer for a fake xlsx
    const blob = new Blob(['fake content'], { type: 'application/octet-stream' });
    formData.append('file', blob, 'test.xlsx');

    try {
        console.log('Sending request...');
        const res = await fetch('http://localhost:3000/api/scan', {
            method: 'POST',
            body: formData
        });

        console.log('Status:', res.status);
        if (res.ok) {
            console.log('Success:', await res.json());
        } else {
            console.log('Fail:', await res.text());
        }
    } catch (e) {
        console.error('Fetch Error:', e.message);
    }
}

testScan();
