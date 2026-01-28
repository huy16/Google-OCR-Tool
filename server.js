const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const multer = require('multer');
const path = require('path');
// Point to the correct updated service in Web_App
const geocoderService = require('./Web_App/geocoder_service');
const fs = require('fs');

// Ensure directories exist
const uploadDir = path.join(process.cwd(), 'uploads');
const outputDir = path.join(process.cwd(), 'outputs');

if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
    console.log('Created uploads directory');
}
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
    console.log('Created outputs directory');
}

// Global Error Handlers to prevent crash
process.on('uncaughtException', (err) => {
    console.error('💥 UNCAUGHT EXCEPTION:', err);
    // Keep process alive if possible, or at least log
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('💥 UNHANDLED REJECTION:', reason);
});

const app = express();
const server = http.createServer(app);
const io = socketIo(server);

// Middleware: Request Logger
app.use((req, res, next) => {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
    next();
});

// Setup Uploads
const upload = multer({ dest: path.join(process.cwd(), 'uploads') }); // Uploads go to real disk
app.use(express.static(path.join(__dirname, 'Web_App', 'public'))); // Assets from inside exe (snapshot)

// API: Upload & Start
app.post('/api/start', (req, res, next) => {
    // Custom wrap for multer to catch errors
    upload.single('file')(req, res, (err) => {
        if (err) {
            console.error('❌ Multer Upload Error:', err);
            return res.status(500).json({ error: 'Upload failed: ' + err.message });
        }
        next();
    });
}, (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    console.log('✅ File uploaded:', req.file.path);

    const jobId = Date.now().toString();
    const inputPath = req.file.path;
    const outputDir = path.join(process.cwd(), 'outputs');

    // Parse options (safely)
    let headlessOption = "new";
    if (req.body.headless === 'false') headlessOption = false;

    const options = {
        projectCodeFilter: req.body.projectCode,
        provinceFilter: req.body.province,
        districtFilter: req.body.district,
        surveyInfoFilter: req.body.surveyInfo,
        headless: headlessOption
    };

    // Return Job ID immediately
    res.json({ jobId, message: 'Job started' });

    // Run Service in Background
    // Wrap in try-catch just in case immediate sync error
    try {
        geocoderService.run(inputPath, outputDir, jobId, options);
    } catch (runErr) {
        console.error('❌ Error triggering run:', runErr);
    }
});

// API: Stop Processing
app.post('/api/stop', (req, res) => {
    console.log('--- API /stop HIT ---');
    geocoderService.stop();
    res.json({ message: 'Stop signal sent' });
});

// API: Scan File for Filters
app.post('/api/scan', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    console.log('API Scan request received for file:', req.file.path);
    if (!fs.existsSync(req.file.path)) {
        console.error('File does not exist on disk:', req.file.path);
        return res.status(500).send('Uploaded file not found on server');
    }

    try {
        const result = await geocoderService.scanFile(req.file.path);
        console.log('--- DEBUG SCAN RESULT ---');
        console.log('Total Rows:', result.totalRows);
        console.log('Keys:', Object.keys(result));
        // res.json(result);
        res.json(result);
    } catch (err) {
        console.error('Scan Error:', err);
        res.status(500).send(err.message || 'Error occurred during scan');
    }
});

// Socket.IO for Real-time Progress
geocoderService.on('log', (msg) => {
    io.emit('log', msg);
});

geocoderService.on('progress', (data) => {
    io.emit('progress', data);
});

geocoderService.on('complete', (result) => {
    io.emit('complete', {
        filename: path.basename(result.file),
        downloadUrl: `/outputs/${path.basename(result.file)}`
    });
});

geocoderService.on('error', (err) => {
    io.emit('error', err);
});

// Serve Outputs
app.use('/outputs', express.static(path.join(process.cwd(), 'outputs')));

// Global Express Error Handler
app.use((err, req, res, next) => {
    console.error('💥 Express Global Error:', err);
    if (!res.headersSent) {
        res.status(500).json({ error: 'Internal Server Error: ' + err.message });
    }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);

    // Auto-open browser (Windows)
    try {
        const { exec } = require('child_process');
        exec(`start http://localhost:${PORT}`);
    } catch (e) {
        console.error('Failed to auto-open browser:', e);
    }
});
