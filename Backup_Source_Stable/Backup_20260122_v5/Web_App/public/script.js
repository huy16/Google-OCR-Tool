
let socket = null;
let currentUploadedFile = null;

// Ensure IO is safely initialized
function initSocket() {
    if (typeof io !== 'undefined') {
        try {
            socket = io();
            console.log('Socket initialized');
            return true;
        } catch (e) {
            console.error('Socket init error:', e);
            return false;
        }
    } else {
        console.warn('Socket.io library not loaded');
        return false;
    }
}

document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM Ready - App v6 (Safe Mode)');

    const hasSocket = initSocket();

    const els = {
        dropZone: document.getElementById('dropZone'),
        fileInput: document.getElementById('fileInput'),
        startBtn: document.getElementById('startBtn'),
        removeFile: document.getElementById('removeFile'),
        uploadCard: document.getElementById('uploadCard'),
        statusCard: document.getElementById('statusCard'),
        resultCard: document.getElementById('resultCard'),
        fileName: document.getElementById('fileName'),
        fileSize: document.getElementById('fileSize'),
        fileInfo: document.getElementById('fileInfo'),
        filterConfig: document.getElementById('filterConfig')
    };

    // 1. File Selection Logic
    els.dropZone.addEventListener('click', () => {
        console.log('Dropzone clicked');
        // alert('Dropzone clicked'); // Optional: Uncomment if console is hard to check
        els.fileInput.value = '';
        try {
            els.fileInput.click();
            console.log('FileInput click triggered');
        } catch (e) {
            alert('Error triggering file input: ' + e.message);
        }
    });

    els.dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        els.dropZone.classList.add('drag-active');
    });

    els.dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        els.dropZone.classList.remove('drag-active');
    });

    els.dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        els.dropZone.classList.remove('drag-active');
        console.log('File dropped');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    els.fileInput.addEventListener('change', (e) => {
        console.log('File changed', e.target.files);
        // alert('File selected: ' + (e.target.files[0] ? e.target.files[0].name : 'None'));
        if (e.target.files.length) handleFile(e.target.files[0]);
    });



    els.removeFile.addEventListener('click', (e) => {
        e.stopPropagation();
        currentUploadedFile = null;
        els.fileInput.value = '';
        els.fileInfo.classList.add('hidden');

        const h3 = els.dropZone.querySelector('h3');
        if (h3) h3.classList.remove('hidden');

        const p = els.dropZone.querySelector('p');
        if (p) p.classList.remove('hidden');

        const icon = els.dropZone.querySelector('.icon-wrapper');
        if (icon) icon.classList.remove('hidden');
        els.filterConfig.classList.add('hidden');
        els.startBtn.disabled = true;
        // Reset button state
        els.startBtn.innerHTML = '<i class="fa-solid fa-play"></i> Start Processing';
        els.startBtn.classList.remove('btn-warning');
        els.startBtn.onclick = null; // Remove force override
    });

    // 2. Handling File & Scanning
    function handleFile(file) {
        console.log('Handle File:', file.name);
        currentUploadedFile = file;

        // Update UI
        if (els.fileName) els.fileName.textContent = file.name;
        if (els.fileSize) els.fileSize.textContent = (file.size / 1024 / 1024).toFixed(2) + ' MB';
        els.fileInfo.classList.remove('hidden');

        const h3 = els.dropZone.querySelector('h3');
        if (h3) h3.classList.add('hidden');

        const p = els.dropZone.querySelector('p');
        if (p) p.classList.add('hidden');

        const icon = els.dropZone.querySelector('.icon-wrapper');
        if (icon) icon.classList.add('hidden');

        // Show Scan UI
        els.filterConfig.classList.remove('hidden');

        // Set Loading
        const inputs = ['projectCodeInput', 'provinceInput', 'districtInput', 'surveyInput'];
        inputs.forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.innerHTML = '<option>Scanning...</option>';
                el.disabled = true;
            }
        });
        els.startBtn.disabled = true;

        // Call API
        scanFileAPI(file);
    }

    // FORCE START Feature
    function showForceStart(file) {
        els.startBtn.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Force Start (Skip Scan)';
        els.startBtn.classList.add('btn-warning');
        els.startBtn.disabled = false;

        els.startBtn.onclick = (e) => {
            e.stopPropagation();
            e.preventDefault();
            // Manually enable dropdowns with defaults
            ['projectCodeInput', 'provinceInput', 'districtInput', 'surveyInput'].forEach(id => {
                const el = document.getElementById(id);
                el.innerHTML = '<option value="">All</option>';
                el.disabled = false;
            });
            // Trigger normal start
            triggerStart(file);
        };
    }

    async function scanFileAPI(file) {
        const formData = new FormData();
        formData.append('file', file);

        // Timeout Protection: Show Force Start after 8s
        const timeout = setTimeout(() => {
            if (els.startBtn.disabled) {
                console.warn('Scan slow, showing force start');
                showForceStart(file);
            }
        }, 8000);

        try {
            console.log('Scanning...');
            const res = await fetch('/api/scan', { method: 'POST', body: formData });
            clearTimeout(timeout);

            if (!res.ok) {
                const txt = await res.text();
                throw new Error(txt);
            }
            const data = await res.json();
            console.log('Scan result:', data);

            // Reset to normal if successful
            els.startBtn.classList.remove('btn-warning');
            els.startBtn.innerHTML = '<i class="fa-solid fa-play"></i> Start Processing';
            els.startBtn.onclick = null; // Remove force override

            populateFilters(data);
        } catch (e) {
            clearTimeout(timeout);
            console.error(e);
            alert('Error scanning: ' + e.message + '\n\nYou can click "Force Start" to proceed anyway.');
            showForceStart(file);
        }
    }

    // Global Filter Data
    let GLOBAL_FILTER_DATA = [];

    function populateFilters(data) {
        try {
            GLOBAL_FILTER_DATA = data.filterData || [];

            // Show Total Rows
            const h3 = document.querySelector('#filterConfig h3');
            if (h3) {
                h3.innerHTML = `<i class="fa-solid fa-filter"></i> Start Configuration <span style="font-size: 14px; font-weight: 400; color: #666; margin-left:10px;">(Found ${data.totalRows} rows)</span>`;
            }

            // Init Event Listeners for Cascading
            setupCascadingListeners();

            // Initial Render (No filters selected)
            renderDropdowns();

            console.log('Enabling Start Button...');
            els.startBtn.disabled = false;
        } catch (e) {
            console.error('Populate filter error:', e);
            alert('Filter display error: ' + e.message + '\nStack: ' + e.stack);
            els.startBtn.disabled = false;
        }
    }

    function setupCascadingListeners() {
        // Use a flag to prevent duplicate listeners if scanFileAPI is called multiple times
        if (window.hasSetupCascadingListeners) return;
        window.hasSetupCascadingListeners = true;

        const provInput = document.getElementById('provinceInput');
        if (provInput) {
            provInput.addEventListener('change', () => {
                // Reset child filters
                const distInput = document.getElementById('districtInput');
                const survInput = document.getElementById('surveyInput');
                if (distInput) distInput.value = '';
                if (survInput) survInput.value = '';
                renderDropdowns();
            });
        }

        const distInput = document.getElementById('districtInput');
        if (distInput) {
            distInput.addEventListener('change', () => {
                const survInput = document.getElementById('surveyInput');
                if (survInput) survInput.value = '';
                renderDropdowns();
            });
        }

        const projInput = document.getElementById('projectCodeInput');
        if (projInput) {
            projInput.addEventListener('change', () => {
                renderDropdowns();
            });
        }

        const survInput = document.getElementById('surveyInput');
        if (survInput) {
            survInput.addEventListener('change', () => {
                renderDropdowns();
            });
        }
    }

    function renderDropdowns() {
        // Get current selections (safely)
        const getVal = (id) => {
            const el = document.getElementById(id);
            return el ? el.value : '';
        };

        const selectedProj = getVal('projectCodeInput');
        const selectedProv = getVal('provinceInput');
        const selectedDist = getVal('districtInput');
        const selectedSurv = getVal('surveyInput');

        // Filter Logic
        const filterData = (criteria) => {
            return GLOBAL_FILTER_DATA.filter(item => {
                if (criteria.proj && item.proj !== criteria.proj) return false;
                if (criteria.prov && item.prov !== criteria.prov) return false;
                if (criteria.dist && item.dist !== criteria.dist) return false;
                return true;
            });
        };

        // 1. Render Projects (Source: All Data)
        updateDropdown('projectCodeInput', filterData({}), 'proj', selectedProj);

        // 2. Render Provinces (Filter by Project)
        updateDropdown('provinceInput', filterData({ proj: selectedProj }), 'prov', selectedProv);

        // 3. Render Districts (Filter by Project + Province)
        updateDropdown('districtInput', filterData({ proj: selectedProj, prov: selectedProv }), 'dist', selectedDist);

        // 4. Render Survey (Filter by Project + Province + District)
        updateDropdown('surveyInput', filterData({ proj: selectedProj, prov: selectedProv, dist: selectedDist }), 'surv', selectedSurv);
    }

    function updateDropdown(id, dataSubset, key, currentSelection) {
        const el = document.getElementById(id);
        if (!el) return;

        // Aggregate unique values and their counts
        const valMap = new Map();
        let totalInSubset = 0;

        dataSubset.forEach(item => {
            const val = item[key];
            if (!val) return;
            valMap.set(val, (valMap.get(val) || 0) + item.count);
            totalInSubset += item.count;
        });

        // Sort alphabetically
        const sorted = [...valMap.entries()].sort((a, b) => String(a[0]).localeCompare(String(b[0])));

        // Build HTML
        let html = `<option value="">All (${totalInSubset})</option>`;

        let selectionIsValid = false;

        sorted.forEach(([val, count]) => {
            const valStr = String(val);
            const safeVal = valStr.replace(/"/g, '&quot;');

            let sel = '';
            if (valStr === currentSelection) {
                sel = 'selected';
                selectionIsValid = true;
            }
            html += `<option value="${safeVal}" ${sel}>${valStr} (${count})</option>`;
        });

        el.innerHTML = html;
        el.disabled = false;

        // If selection invalid, reset to empty (All)
        if (currentSelection !== '' && !selectionIsValid) {
            el.value = '';
        } else if (currentSelection !== '') {
            el.value = currentSelection;
        }
    }

    // 3. Start Processing Logic
    async function triggerStart(file) {
        addLog('Start button clicked', 'info');
        console.log('Triggering Start...');

        els.startBtn.disabled = true;
        els.startBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';
        els.startBtn.onclick = null; // Clear override

        const formData = new FormData();
        formData.append('file', file);
        formData.append('projectCode', document.getElementById('projectCodeInput').value || ''); // Safe access logic
        formData.append('province', document.getElementById('provinceInput').value || '');
        formData.append('district', document.getElementById('districtInput').value || '');
        formData.append('surveyInfo', document.getElementById('surveyInput').value || '');
        // Checkbox: checked = show browser (headless: false), unchecked = headless: true
        const showBrowser = document.getElementById('headlessToggle').checked;
        formData.append('headless', !showBrowser);

        try {
            console.log('Sending start request...');
            const res = await fetch('/api/start', { method: 'POST', body: formData });
            if (res.ok) {
                // Success UI transition
                console.log('Start Success');
                els.uploadCard.classList.add('hidden');
                document.getElementById('statusCard').classList.remove('hidden');
                addLog('System started...', 'system');
            } else {
                const txt = await res.text();
                console.error('Start Failed:', txt);
                alert('Start failed (Server Error): ' + txt);
                resetBtnState();
            }
        } catch (e) {
            console.error('Start Network Error:', e);
            alert('Network error: ' + e.message);
            resetBtnState();
        }
    }

    function resetBtnState() {
        els.startBtn.disabled = false;
        els.startBtn.innerHTML = '<i class="fa-solid fa-play"></i> Start Processing';
    }

    // Default Listener (Normal flow)
    els.startBtn.addEventListener('click', async () => {
        if (!currentUploadedFile) {
            console.warn('Click ignored: No file');
            return;
        }
        if (!els.startBtn.classList.contains('btn-warning')) {
            triggerStart(currentUploadedFile);
        }
    });

    // 4. Socket Listeners
    if (hasSocket) {
        socket.on('log', (msg) => {
            let type = 'info';
            if (msg.toLowerCase().includes('error')) type = 'error';
            else if (msg.includes('✅')) type = 'success';
            addLog(msg, type);
        });

        socket.on('progress', (data) => {
            const bar = document.getElementById('progressBar');
            const txt = document.getElementById('progressText');
            const detail = document.getElementById('statusDetail');
            if (bar) bar.style.width = data.percent + '%';
            if (txt) txt.textContent = data.percent + '%';
            if (detail) {
                // Check if we have extended data
                if (data.total) {
                    detail.textContent = `[${data.current}/${data.total}] Row ${data.row}: ${data.message}`;
                } else {
                    detail.textContent = `Row ${data.row}: ${data.message}`;
                }
            }
        });

        const stopBtn = document.getElementById('stopBtn');
        if (stopBtn) {
            stopBtn.addEventListener('click', async () => {
                if (confirm('Bạn có chắc muốn dừng không? Kết quả đã chạy sẽ được lưu lại.')) {
                    stopBtn.disabled = true;
                    stopBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Stopping...';
                    try {
                        await fetch('/api/stop', { method: 'POST' });
                    } catch (e) {
                        console.error('Stop error:', e);
                    }
                }
            });
        }

        socket.on('complete', (result) => {
            addLog('Completed!', 'success');
            setTimeout(() => {
                document.getElementById('statusCard').classList.add('hidden');
                const resCard = document.getElementById('resultCard');
                resCard.classList.remove('hidden');
                document.getElementById('downloadLink').href = result.downloadUrl;
                document.getElementById('totalProcessed').textContent = result.processed || 'Done';
            }, 1500);
        });

        socket.on('error', (err) => {
            addLog(`Error: ${err}`, 'error');
        });
    }

    function addLog(msg, type = 'info') {
        const div = document.createElement('div');
        div.className = `log-line ${type}`;
        div.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
        const body = document.getElementById('logBody');
        if (body) {
            body.appendChild(div);
            body.scrollTop = body.scrollHeight;
        }
    }
});
