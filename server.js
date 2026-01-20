const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const multer = require('multer');
const path = require('path');
const geocoderService = require('./geocoder_service');
const fs = require('fs');

// Ensure directories exist
const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'outputs');

if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
    console.log('Created uploads directory');
}
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
    console.log('Created outputs directory');
}

const app = express();
const server = http.createServer(app);
const io = socketIo(server);

// Setup Uploads
const upload = multer({ dest: 'uploads/' });
app.use(express.static('public'));

// API: Upload & Start
app.post('/api/start', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    const jobId = Date.now().toString();
    const inputPath = req.file.path;
    const outputDir = path.join(__dirname, 'outputs');
    const options = {
        projectCodeFilter: req.body.projectCode,
        provinceFilter: req.body.province,
        districtFilter: req.body.district,
        surveyInfoFilter: req.body.surveyInfo
    };

    // Return Job ID immediately
    res.json({ jobId, message: 'Job started' });

    // Run Service in Background
    geocoderService.run(inputPath, outputDir, jobId, options);
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
        // Clean up temp file? Maybe keep for next step? 
        // For simple flow, user re-uploads or we keep it. 
        // Ideally we keep it and return a session ID, but here explicit re-upload is safer for stateless.
        // Actually script.js re-uploads on scan? Ideally we just scan temp file.
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
app.use('/outputs', express.static('outputs'));

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
