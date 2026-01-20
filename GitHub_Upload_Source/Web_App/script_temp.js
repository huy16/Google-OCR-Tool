document.addEventListener('DOMContentLoaded', () => {
    // Basic elements
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const startBtn = document.getElementById('startBtn');

    // Debug check
    if (!dropZone || !fileInput) {
        console.error('Missing elements');
        return;
    }

    // --- File Selection ---
    dropZone.addEventListener('click', () => {
        fileInput.value = ''; // Reset
        fileInput.click();
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-active');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-active');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-active');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    function handleFile(file) {
        // UI Update
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileInfo').classList.remove('hidden');
        dropZone.querySelector('h3').classList.add('hidden');
        dropZone.querySelector('p').classList.add('hidden');
        dropZone.querySelector('.icon-wrapper').classList.add('hidden');

        // Reset Filters UI
        document.getElementById('filterConfig').classList.remove('hidden');
        ['projectCodeInput', 'provinceInput', 'districtInput', 'surveyInput'].forEach(id => {
            const el = document.getElementById(id);
            el.innerHTML = '<option>Scanning...</option>';
            el.disabled = true;
        });
        startBtn.disabled = true;

        // Start Scan
        scanFileAPI(file);
    }

    async function scanFileAPI(file) {
        const formData = new FormData();
        formData.append('file', file);

        try {
            const res = await fetch('/api/scan', { method: 'POST', body: formData });
            if (!res.ok) {
                const txt = await res.text();
                throw new Error(txt);
            }
            const data = await res.json();
            populateDropdowns(data);
        } catch (e) {
            console.error(e);
            alert('Error scanning file: ' + e.message);
            // Reset UI slightly
            document.getElementById('fileName').textContent = 'Error scanning';
        }
    }

    function populateDropdowns(data) {
        const fill = (id, items) => {
            const el = document.getElementById(id);
            el.innerHTML = '<option value="">All</option>';
            items.forEach(i => el.innerHTML += `<option value="${i}">${i}</option>`);
            el.disabled = false;
        };

        fill('projectCodeInput', data.projects);
        fill('provinceInput', data.provinces);
        fill('districtInput', data.districts);
        fill('surveyInput', data.surveys);

        // Auto select defaults
        const setVal = (id, val) => {
            const el = document.getElementById(id);
            for (let opt of el.options) {
                if (opt.value.includes(val)) {
                    el.value = opt.value;
                    break;
                }
            }
        };
        setVal('projectCodeInput', '2026_bidding');
        setVal('provinceInput', 'Hồ Chí Minh');

        startBtn.disabled = false;
    }

    // --- Start Processing ---
    startBtn.addEventListener('click', async () => {
        // We need the file again. Since we didn't sync input for Drop, we use fileInput for Click. 
        // Or store file globally?
        // Simplest: Check fileInput. If empty (DragDrop), alert user or store file in variable.
        // Let's store file in a variable

        // Wait, for FormData upload in Start, we need the file object.
        // If drag & drop, fileInput is empty.
        // Let's rely on re-uploading via handleFile? 
        // No, 'Start' calls '/api/start'.

        // BETTER: When handleFile runs, ASSIGN to fileInput if possible, or store in closure.
        // We cannot easily assign to fileInput if drag-dropped.
        // Solution: Store the current file object in a global var.
    });
});

// We need to define `currentFile` outside to access it
let currentFile = null;

// Re-attach listener correctly with logic inside
document.addEventListener('DOMContentLoaded', () => {
    // ... same setup ...
    const startBtn = document.getElementById('startBtn');

    // Override handleFile to store file
    const originalHandle = (file) => {
        currentFile = file;
        // ... (UI updates from above)
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileInfo').classList.remove('hidden');
        document.getElementById('dropZone').querySelector('h3').classList.add('hidden');
        document.getElementById('dropZone').querySelector('p').classList.add('hidden');
        document.getElementById('dropZone').querySelector('.icon-wrapper').classList.add('hidden');

        document.getElementById('filterConfig').classList.remove('hidden');
        ['projectCodeInput', 'provinceInput', 'districtInput', 'surveyInput'].forEach(id => {
            const el = document.getElementById(id);
            el.innerHTML = '<option>Scanning...</option>';
            el.disabled = true;
        });
        startBtn.disabled = true;

        scanFileAPI(file);
    };

    // ... event listeners using originalHandle ...
    // Using simple global function references for clarity in rewrites
});

// Let's write the full clean file content directly
