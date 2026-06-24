/* ═══════════════════════════════════════════════════════════
   AttendQR v3 — app.js (100% Firebase - NO IndexedDB)
   All data from Firebase Firestore only
═══════════════════════════════════════════════════════════ */

// ════════════════════════════════════════════
// 🔥 FIREBASE CONFIG
// ════════════════════════════════════════════
const firebaseConfig = {
    apiKey: "AIzaSyAtw1h7A3DkcJwv5BEiif2RiwBCDCIuHxQ",
    authDomain: "attendance-16ed5.firebaseapp.com",
    projectId: "attendance-16ed5",
    storageBucket: "attendance-16ed5.firebasestorage.app",
    messagingSenderId: "252998999489",
    appId: "1:252998999489:web:a64af8feb6b6786a602340",
    measurementId: "G-9VPCEE4H91"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const firestore = firebase.firestore();


/* ── Settings ────────────────────────────────────────────── */

// Get settings from Firebase
async function getSettings() {
    try {
        const settingsDoc = await firestore.collection('settings').doc('appSettings').get();
        if (settingsDoc.exists) {
            return settingsDoc.data();
        }
        return { scannerUrl: '' };
    } catch (err) {
        console.warn('Error loading settings:', err);
        return { scannerUrl: '' };
    }
}

// Save settings to Firebase
async function saveSettingsToFirebase(scannerUrl) {
    try {
        await firestore.collection('settings').doc('appSettings').set({
            scannerUrl: scannerUrl,
            updatedAt: new Date().toISOString()
        });
        console.log('✅ Settings saved to Firebase');
        return true;
    } catch (err) {
        console.error('Error saving settings:', err);
        return false;
    }
}

/* ── Load Settings ────────────────────────────────────────── */
async function loadSettings() {
    const settings = await getSettings();
    document.getElementById('scanner-url').value = settings.scannerUrl || '';
}

/* ── Save Settings ────────────────────────────────────────── */
async function saveSettings() {
    const scannerUrl = document.getElementById('scanner-url').value.trim();

    const success = await saveSettingsToFirebase(scannerUrl);

    const msg = document.getElementById('save-msg');
    if (success) {
        msg.style.display = 'flex';
        setTimeout(() => (msg.style.display = 'none'), 2200);
        await renderEmployees();
        await renderDashboard();
        await renderLog();
    } else {
        alert('Failed to save settings. Please check your connection.');
    }
}

/* ── Employee CRUD Operations ─────────────────────────────── */

// Get ALL employees from Firebase
async function getEmployees() {
    try {
        console.log('📊 Fetching employees from Firebase...');
        const snapshot = await firestore.collection('employees').get();
        let employees = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            employees.push({
                id: doc.id,
                name: data.name || 'Unknown',
                dept: data.dept || data.department || ''
            });
            console.log('📄 Employee:', data.name, '| Dept:', data.dept || data.department || 'None');
        });
        console.log(`✅ Total employees in Firebase: ${employees.length}`);
        return employees;
    } catch (err) {
        console.error('Error fetching employees:', err);
        return [];
    }
}

// Add Employee
async function addEmployee() {
    const name = document.getElementById('emp-name').value.trim();
    const dept = document.getElementById('emp-dept').value.trim();
    if (!name) {
        document.getElementById('emp-name').focus();
        return;
    }

    try {
        await firestore.collection('employees').add({
            name: name,
            dept: dept || '',
            createdAt: new Date().toISOString()
        });

        console.log(`✅ Added employee: ${name}`);

        document.getElementById('emp-name').value = '';
        document.getElementById('emp-dept').value = '';
        document.getElementById('emp-name').focus();

        await renderEmployees();
        await renderDashboard();
    } catch (err) {
        console.error('Error adding employee:', err);
        alert('Failed to add employee: ' + err.message);
    }
}

// Delete Employee
async function removeEmployee(docId) {
    if (!confirm('Remove this employee? Their past attendance records will remain.')) return;

    try {
        await firestore.collection('employees').doc(docId).delete();
        console.log(`✅ Deleted employee`);
        await renderEmployees();
        await renderDashboard();
    } catch (err) {
        console.error('Error deleting employee:', err);
        alert('Failed to delete employee: ' + err.message);
    }
}

// Render employees list
async function renderEmployees() {
    const employees = await getEmployees();
    document.getElementById('emp-count').textContent = employees.length;

    const list = document.getElementById('emp-list');
    if (!employees.length) {
        list.innerHTML = '<div class="empty-state">No employees yet. Add one above or import from Excel.</div>';
        return;
    }
    list.innerHTML = employees.map(e => `
        <div class="emp-row">
            <div class="avatar">${initials(e.name)}</div>
            <div class="emp-info">
                <div class="emp-name">${e.name}</div>
                <div class="emp-meta">${e.dept || 'No department'}</div>
            </div>
            <button class="btn-icon" onclick="removeEmployee('${e.id}')" aria-label="Remove ${e.name}">
                <i class="ti ti-trash"></i>
            </button>
        </div>
    `).join('');
}

// Import employees from Excel
async function importEmployees(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async e => {
        try {
            const wb = XLSX.read(e.target.result, { type: 'binary' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

            if (!rows.length) { alert('No data found in the file.'); return; }

            const normalize = str => String(str).trim().toLowerCase();
            const nameKey = Object.keys(rows[0]).find(k => normalize(k) === 'name');
            const deptKey = Object.keys(rows[0]).find(k => ['department', 'dept'].includes(normalize(k)));

            if (!nameKey) { alert('Could not find a "Name" column.'); return; }

            let added = 0;
            for (const row of rows) {
                const name = String(row[nameKey] || '').trim();
                const dept = deptKey ? String(row[deptKey] || '').trim() : '';
                if (name) {
                    await firestore.collection('employees').add({
                        name: name,
                        dept: dept,
                        createdAt: new Date().toISOString()
                    });
                    added++;
                }
            }

            alert(`Imported ${added} employee${added !== 1 ? 's' : ''} successfully!`);
            await renderEmployees();
            await renderDashboard();
        } catch (err) {
            alert('Failed to read Excel file: ' + err.message);
        }
    };
    reader.readAsBinaryString(file);
    event.target.value = '';
}

/* ── Attendance Operations ───────────────────────────────── */

// Get attendance records from Firebase
async function getAttendance() {
    try {
        console.log('📊 Fetching attendance from Firebase...');
        const snapshot = await firestore.collection('attendance')
            .orderBy('timestamp', 'desc')
            .limit(500)
            .get();

        let records = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            records.push({
                id: doc.id,
                name: data.name || '',
                dept: data.department || data.dept || '',
                date: data.date || '',
                dateISO: data.dateISO || '',
                time: data.time || '',
                device: data.device || '',
                status: data.status || 'Present',
                company: data.company || '',
                location: data.location || '',
                totalTime: data.totalTime || '',
                totalHours: data.totalHours || 0,
                logoutTime: data.logoutTime || ''
            });
        });
        
        console.log(`✅ Total attendance records in Firebase: ${records.length}`);
        return records;
    } catch (err) {
        console.warn('Error fetching attendance:', err);
        return [];
    }
}

/* ── Dashboard ───────────────────────────────────────────── */
async function renderDashboard() {
    const employees = await getEmployees();
    const allAttendance = await getAttendance();
    
    const today = new Date();
    const todayStr = today.toLocaleDateString('en-IN');
    
    console.log('📅 Today:', todayStr);
    console.log('📊 All attendance records:', allAttendance);
    
    const todayLogs = allAttendance.filter(r => r.date === todayStr);

    const total = employees.length;
    const present = todayLogs.length;
    const absentN = Math.max(0, total - present);

    document.getElementById('dash-total').textContent = total;
    document.getElementById('dash-present').textContent = present;
    document.getElementById('dash-absent').textContent = absentN;
    document.getElementById('today-date').textContent = new Date().toLocaleDateString('en-IN', { 
        weekday: 'long',
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
    });

    const presentNames = new Set(todayLogs.map(r => r.name));
    const absentEmps = employees.filter(e => !presentNames.has(e.name));
    const absentEl = document.getElementById('absent-list');

    if (!absentEmps.length) {
        absentEl.innerHTML = '<div class="empty-state" style="color:var(--green)"><i class="ti ti-circle-check"></i> Everyone is present today!</div>';
    } else {
        absentEl.innerHTML = absentEmps.map(e => `
            <div class="emp-row">
                <div class="avatar">${initials(e.name)}</div>
                <div class="emp-info">
                    <div class="emp-name">${e.name}</div>
                    <div class="emp-meta">${e.dept || 'No department'}</div>
                </div>
                <span class="badge badge-red">Absent</span>
            </div>
        `).join('');
    }

    const todayEl = document.getElementById('today-log-list');
    if (!todayLogs.length) {
        todayEl.innerHTML = '<div class="empty-state">No attendance marked yet today.</div>';
    } else {
        todayEl.innerHTML = `
          <div class="log-row log-row-timer log-header">
            <span>Employee</span><span>Dept</span><span>In Time</span><span>Total Time</span><span>Status</span><span>Break</span>
          </div>
        ` + todayLogs.map(r => {
            const safeName = r.name.replace(/'/g, "\\'");
            const hasLogout = r.totalTime && r.logoutTime;
            return `
            <div class="log-row log-row-timer" id="row-${r.name.replace(/\s+/g,'_')}">
                <span style="font-weight:600">${r.name}</span>
                <span style="color:var(--text-secondary)">${r.dept || '—'}</span>
                <span style="color:var(--text-secondary)">${r.time || 'N/A'}</span>
                <span class="timer-cell" id="timer-${r.name.replace(/\s+/g,'_')}">
                  ${hasLogout
                    ? `<span class="timer-display" style="color:var(--green)">${r.totalTime}</span>
                       <span style="font-size:11px;color:var(--text-hint);display:block">out ${r.logoutTime}</span>`
                    : `<span class="timer-display">00:00:00</span>`
                  }
                </span>
                <span><span class="badge ${hasLogout ? 'badge-amber' : 'badge-green'}" id="status-badge-${r.name.replace(/\s+/g,'_')}">${hasLogout ? 'Logged Out' : (r.status || 'Present')}</span></span>
                <span>
                  ${hasLogout
                    ? `<span style="font-size:12px;color:var(--text-hint)">Shift ended</span>`
                    : `<button class="btn-break" id="break-btn-${r.name.replace(/\s+/g,'_')}"
                        onclick="toggleBreak('${safeName}')">
                        <i class="ti ti-coffee"></i> Break
                      </button>`
                  }
                </span>
            </div>`;
        }).join('');
        startDashboardTimers(todayLogs);
    }
}

/* ── QR Generation ───────────────────────────────────────── */
async function generateQR() {
    const settings = await getSettings();
    const scannerUrl = settings.scannerUrl || document.getElementById('scanner-url')?.value?.trim() || '';
    const employees = await getEmployees();

    const urlWarn = document.getElementById('scanner-url-warning');
    const empWarn = document.getElementById('no-employees-warning');
    urlWarn.style.display = 'none';
    empWarn.style.display = 'none';

    if (!scannerUrl) { urlWarn.style.display = 'flex'; return; }
    if (!employees.length) { empWarn.style.display = 'flex'; return; }

    const url = scannerUrl;
    const size = parseInt(document.getElementById('qr-size').value, 10);

    const canvas = document.getElementById('qr-canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, size, size);

    document.getElementById('qr-placeholder').style.display = 'none';
    canvas.style.display = 'block';

    const tempDiv = document.createElement('div');
    new QRCode(tempDiv, {
        text: url,
        width: size,
        height: size,
        correctLevel: QRCode.CorrectLevel.M,
    });

    setTimeout(() => {
        const img = tempDiv.querySelector('img');
        if (img) {
            const i = new Image();
            i.onload = () => ctx.drawImage(i, 0, 0, size, size);
            i.src = img.src;
        }
    }, 260);

    const label = document.getElementById('qr-label');
    label.textContent = scannerUrl;
    label.style.display = 'block';

    document.getElementById('dl-btn').style.display = 'inline-flex';
    document.getElementById('qr-success').style.display = 'flex';
}

function downloadQR() {
    const canvas = document.getElementById('qr-canvas');
    const a = document.createElement('a');
    a.download = 'attendance-qr.png';
    a.href = canvas.toDataURL('image/png');
    a.click();
}

/* ── Attendance Log ──────────────────────────────────────── */
async function renderLog() {
    const search     = (document.getElementById('log-search')?.value || '').toLowerCase();
    const filterDate = document.getElementById('log-date')?.value || '';
    const filterStat = document.getElementById('log-status')?.value || '';  // '', 'Present', 'Absent'

    const allAttendance = await getAttendance();
    const employees     = await getEmployees();
    const today         = todayStr();

    // ── Determine the reference date for absent calculation ──
    // If a date filter is active use that, otherwise default to today
    const refDateISO  = filterDate || today;
    // Convert ISO (YYYY-MM-DD) → en-IN display string to match stored r.date
    const refDateDisp = new Date(refDateISO + 'T00:00:00')
        .toLocaleDateString('en-IN');

    // ── Set of employees who are present on refDate ──
    const presentOnRef = new Set(
        allAttendance
            .filter(r => r.date === refDateDisp || r.dateISO === refDateISO)
            .map(r => r.name)
    );

    // ── Build absent rows from employee list ──
    const absentRows = employees
        .filter(e => !presentOnRef.has(e.name))
        .map(e => ({
            name:    e.name,
            dept:    e.dept || '',
            date:    refDateDisp,
            dateISO: refDateISO,
            time:    '—',
            status:  'Absent',
        }));

    // ── Decide what to show based on filterStat ──
    let rows = [];
    if (filterStat === 'Absent') {
        // Only absent employees for the reference date
        rows = absentRows;
    } else if (filterStat === 'Present') {
        // Only present attendance records
        rows = allAttendance.filter(r =>
            (!filterDate || r.date === refDateDisp || r.dateISO === refDateISO)
        );
    } else {
        // All — present records (all dates) + absent rows for the ref date
        rows = [
            ...allAttendance.filter(r =>
                (!filterDate || r.date === refDateDisp || r.dateISO === refDateISO)
            ),
            ...absentRows,
        ];
    }

    // ── Apply search filter ──
    if (search) {
        rows = rows.filter(r =>
            r.name.toLowerCase().includes(search) ||
            (r.dept || '').toLowerCase().includes(search)
        );
    }

    // ── Stats ──
    const todayPresentCount = allAttendance.filter(r =>
        r.date === new Date(today + 'T00:00:00').toLocaleDateString('en-IN') || r.dateISO === today
    ).length;

    document.getElementById('log-stat-records').textContent = rows.length;
    document.getElementById('log-stat-today').textContent   = todayPresentCount;
    document.getElementById('log-stat-total').textContent   = employees.length;

    // ── Render ──
    const list = document.getElementById('log-list');
    if (!rows.length) {
        list.innerHTML = '<div class="empty-state">No records match your filter.</div>';
        return;
    }

    list.innerHTML = rows.map(r => {
        const isAbsent   = r.status === 'Absent';
        const badgeCls   = isAbsent ? 'badge-red' : 'badge-green';
        const viewBtn    = isAbsent
            ? `<span style="font-size:12px;color:var(--text-hint)">—</span>`
            : `<button class="btn-icon" onclick="viewAttendance('${r.name.replace(/'/g,"\\'")}', '${r.date}')" style="padding:4px 10px;">
                   <i class="ti ti-eye"></i> View
               </button>`;
        // Total time display
        let totalTimeDisplay = '—';
        if (!isAbsent) {
            if (r.totalTime) {
                totalTimeDisplay = `<span style="font-family:monospace;font-weight:700;color:var(--green)">${r.totalTime}</span>`;
            } else if (r.time) {
                totalTimeDisplay = `<span style="color:var(--text-hint);font-size:12px">In progress…</span>`;
            }
        }
        return `
        <div class="log-row-full">
            <span style="font-weight:600">${r.name}</span>
            <span style="color:var(--text-secondary)">${r.dept || '—'}</span>
            <span style="color:var(--text-secondary)">${r.date || r.dateISO}</span>
            <span style="color:var(--text-secondary)">${r.time || '—'}</span>
            <span>${totalTimeDisplay}</span>
            <span><span class="badge ${badgeCls}">${r.status}</span></span>
            ${viewBtn}
        </div>`;
    }).join('');
}

// ── View Single Employee Attendance ──
async function viewAttendance(name, date) {
    const all = await getAttendance();
    const records = all.filter(r => r.name === name && (r.date === date || r.dateISO === date));

    if (!records || records.length === 0) {
        alert(`No attendance record found for ${name} on ${date}`);
        return;
    }

    const record = records[0];
    const message = `
📋 Attendance Details
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
👤 Employee: ${record.name}
🏢 Department: ${record.dept || 'N/A'}
📅 Date: ${record.date || record.dateISO}
🕐 Time In: ${record.time || 'N/A'}
🕐 Logout: ${record.logoutTime || 'Not logged out'}
⏱ Total Time: ${record.totalTime || 'In progress'}
📱 Device: ${record.device || 'N/A'}
📊 Status: ${record.status || 'Present'}
🏢 Company: ${record.company || 'N/A'}
📍 Location: ${record.location || 'N/A'}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  `;

    alert(message);
}

async function clearLog() {
    if (!confirm('Delete ALL attendance records? This cannot be undone.')) return;

    try {
        const all = await getAttendance();
        for (const record of all) {
            await firestore.collection('attendance').doc(record.id).delete();
        }
        await renderLog();
        await renderDashboard();
    } catch (err) {
        alert('Error clearing records: ' + err.message);
    }
}

async function exportXLSX() {
    const all = await getAttendance();
    if (!all.length) { alert('No attendance records to export.'); return; }

    const headers = ['Name', 'Department', 'Date', 'Time In', 'Logout', 'Total Time', 'Device', 'Status', 'Company', 'Location'];
    const rows = all.map(r => [
        r.name || '',
        r.dept || '',
        r.date || r.dateISO || '',
        r.time || '',
        r.logoutTime || '',
        r.totalTime || '',
        r.device || '',
        r.status || 'Present',
        r.company || '',
        r.location || '',
    ]);

    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = [20, 18, 12, 10, 30, 10, 18, 18].map(w => ({ wch: w }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Attendance');
    XLSX.writeFile(wb, 'attendance-' + new Date().toISOString().slice(0, 10) + '.xlsx');
}

/* ── Tab navigation ──────────────────────────────────────── */
function showTab(tabId) {
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    document.getElementById('tab-' + tabId).classList.add('active');
    const navBtn = document.querySelector(`.nav-item[data-tab="${tabId}"]`);
    if (navBtn) navBtn.classList.add('active');

    if (tabId === 'dashboard') renderDashboard();
    if (tabId === 'employees') renderEmployees();
    if (tabId === 'log') renderLog();
    if (tabId === 'economy') renderEconomy();
    if (tabId === 'billing') renderBillingTable();
}

/* ══════════════════════════════════════════════════════
   THEME PANEL — multi-theme + background image
══════════════════════════════════════════════════════ */

const GRADIENTS = {
  gradient1: 'linear-gradient(135deg,#667eea,#764ba2)',
  gradient2: 'linear-gradient(135deg,#f093fb,#f5576c)',
  gradient3: 'linear-gradient(135deg,#4facfe,#00f2fe)',
  gradient4: 'linear-gradient(135deg,#43e97b,#38f9d7)',
  gradient5: 'linear-gradient(135deg,#fa709a,#fee140)',
};

function openThemePanel() {
  document.getElementById('theme-panel').classList.add('open');
  document.getElementById('theme-panel-overlay').classList.add('visible');
}
function closeThemePanel() {
  document.getElementById('theme-panel').classList.remove('open');
  document.getElementById('theme-panel-overlay').classList.remove('visible');
}

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('attendqr_theme', theme);
  // Update active swatch
  document.querySelectorAll('.swatch').forEach(s => s.classList.remove('active'));
  const target = document.querySelector('.swatch-' + theme);
  if (target) target.classList.add('active');
}

function setBg(key) {
  const body = document.body;
  const previewWrap = document.getElementById('bg-preview-wrap');

  // Clear active state on all preset buttons
  document.querySelectorAll('.bg-preset').forEach(b => b.classList.remove('active'));

  if (key === 'none') {
    body.style.backgroundImage = '';
    body.classList.remove('has-bg-image');
    localStorage.removeItem('attendqr_bg');
    if (previewWrap) previewWrap.style.display = 'none';
    const noneBtn = document.querySelector('.bg-none');
    if (noneBtn) noneBtn.classList.add('active');
    return;
  }

  const isGradient = GRADIENTS[key];
  if (isGradient) {
    body.style.backgroundImage = isGradient;
    body.classList.add('has-bg-image');
    localStorage.setItem('attendqr_bg', JSON.stringify({ type: 'gradient', value: isGradient }));
    if (previewWrap) previewWrap.style.display = 'none';
    // Mark active preset
    const presets = document.querySelectorAll('.bg-preset:not(.bg-none)');
    const idx = Object.keys(GRADIENTS).indexOf(key);
    if (presets[idx]) presets[idx].classList.add('active');
  }
}

function handleBgUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const dataUrl = e.target.result;
    document.body.style.backgroundImage = `url(${dataUrl})`;
    document.body.classList.add('has-bg-image');
    // Store in localStorage (may fail for very large images, graceful fallback)
    try {
      localStorage.setItem('attendqr_bg', JSON.stringify({ type: 'image', value: dataUrl }));
    } catch(err) { /* image too large for localStorage — still applied in-session */ }
    const preview = document.getElementById('bg-preview-img');
    const previewWrap = document.getElementById('bg-preview-wrap');
    if (preview && previewWrap) {
      preview.src = dataUrl;
      previewWrap.style.display = 'block';
    }
    document.querySelectorAll('.bg-preset').forEach(b => b.classList.remove('active'));
  };
  reader.readAsDataURL(file);
  event.target.value = '';
}

function setAccent(color, dark, light) {
  document.documentElement.style.setProperty('--accent', color);
  document.documentElement.style.setProperty('--accent-dark', dark);
  document.documentElement.style.setProperty('--accent-light', light);
  document.documentElement.style.setProperty('--blue', color);
  document.documentElement.style.setProperty('--blue-dark', dark);
  document.documentElement.style.setProperty('--blue-light', light);
  localStorage.setItem('attendqr_accent', JSON.stringify({ color, dark, light }));
}

function applyStoredTheme() {
  // Color theme
  const stored = localStorage.getItem('attendqr_theme') || 'light';
  document.documentElement.setAttribute('data-theme', stored);
  const sw = document.querySelector('.swatch-' + stored);
  if (sw) sw.classList.add('active');

  // Accent
  try {
    const accent = JSON.parse(localStorage.getItem('attendqr_accent') || 'null');
    if (accent) setAccent(accent.color, accent.dark, accent.light);
  } catch(e) {}

  // Background
  try {
    const bg = JSON.parse(localStorage.getItem('attendqr_bg') || 'null');
    if (bg) {
      if (bg.type === 'gradient') {
        document.body.style.backgroundImage = bg.value;
        document.body.classList.add('has-bg-image');
      } else if (bg.type === 'image') {
        document.body.style.backgroundImage = `url(${bg.value})`;
        document.body.classList.add('has-bg-image');
        const preview = document.getElementById('bg-preview-img');
        const previewWrap = document.getElementById('bg-preview-wrap');
        if (preview && previewWrap) { preview.src = bg.value; previewWrap.style.display = 'block'; }
      }
    }
  } catch(e) {}
}

// Legacy toggleTheme kept for safety (not exposed in new UI)
function toggleTheme() { setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'); }

/* ══════════════════════════════════════════════════════
   STOCKS COMING SOON MODAL
══════════════════════════════════════════════════════ */
function showStocksModal() {
  document.getElementById('stocks-modal').style.display = 'block';
  document.getElementById('stocks-modal-overlay').style.display = 'block';
  closeSidebar();
}
function closeStocksModal() {
  document.getElementById('stocks-modal').style.display = 'none';
  document.getElementById('stocks-modal-overlay').style.display = 'none';
}

/* ── Mobile sidebar ──────────────────────────────────────── */
function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
    document.getElementById('sidebar-overlay').classList.toggle('visible');
}

function closeSidebar() {
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('sidebar-overlay').classList.remove('visible');
}

/* ── Utilities ───────────────────────────────────────────── */
function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function initials(name) {
    return (name || '?').split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
}

/* ── Quick Mark Attendance ── */
async function quickMarkAttendance() {
    if (!confirm('Mark yourself as Present right now?')) return;
    
    try {
        const employees = await getEmployees();
        if (employees.length === 0) {
            alert('No employees found. Please add employees first.');
            return;
        }
        
        const emp = employees[0];
        const now = new Date();
        const date = now.toLocaleDateString('en-IN');
        const time = now.toLocaleTimeString('en-IN');
        const todayStr = now.toISOString().split('T')[0];
        
        const allAttendance = await getAttendance();
        const alreadyMarked = allAttendance.some(r => r.name === emp.name && r.date === date);
        
        if (alreadyMarked) {
            alert('You have already marked attendance today!');
            return;
        }
        
        await firestore.collection('attendance').add({
            name: emp.name,
            department: emp.dept || '',
            date: date,
            dateISO: todayStr,
            time: time,
            device: 'Quick Mark',
            status: 'Present',
            company: 'CityWare',
            location: 'Head Office',
            timestamp: firebase.firestore.FieldValue.serverTimestamp(),
            createdAt: new Date().toISOString()
        });
        
        alert('✅ Attendance marked successfully!');
        renderDashboard();
        renderLog();
    } catch (err) {
        alert('Error: ' + err.message);
    }
}

/* ── Force Refresh from Firebase ── */
async function forceRefresh() {
    console.log('🔄 Force refreshing from Firebase...');
    await renderEmployees();
    await renderDashboard();
    await renderLog();
    console.log('✅ Refresh complete!');
}

// Make functions globally available
window.forceRefresh = forceRefresh;
window.quickMarkAttendance = quickMarkAttendance;
window.addEmployee = addEmployee;
window.removeEmployee = removeEmployee;
window.importEmployees = importEmployees;
window.saveSettings = saveSettings;
window.generateQR = generateQR;
window.downloadQR = downloadQR;
window.exportXLSX = exportXLSX;
window.clearLog = clearLog;
window.viewAttendance = viewAttendance;
window.showTab = showTab;
window.toggleTheme = toggleTheme;
window.toggleSidebar = toggleSidebar;
window.closeSidebar = closeSidebar;
// Theme panel
window.openThemePanel  = openThemePanel;
window.closeThemePanel = closeThemePanel;
window.setTheme        = setTheme;
window.setBg           = setBg;
window.handleBgUpload  = handleBgUpload;
window.setAccent       = setAccent;
// Stocks modal
window.showStocksModal  = showStocksModal;
window.closeStocksModal = closeStocksModal;

/* ── Dashboard Live Timers ───────────────────────────────── */
let _dashTimerInterval = null;
// breakStates: { [name]: { onBreak: bool, breakStart: Date|null, totalBreakMs: number } }
let _breakStates = {};
// checkinTimes: { [name]: Date } — when they first marked attendance
let _checkinTimes = {};

function startDashboardTimers(todayLogs) {
    if (_dashTimerInterval) clearInterval(_dashTimerInterval);

    // Only tick timers for employees who haven't logged out yet
    const activeLogs = todayLogs.filter(r => !r.totalTime || !r.logoutTime);

    // Build checkin times from log (time string like "3:45:22 pm")
    activeLogs.forEach(r => {
        const key = r.name;
        if (!_checkinTimes[key]) {
            // Parse time from record; fall back to now if unparseable
            const parsed = r.time ? new Date(new Date().toDateString() + ' ' + r.time) : new Date();
            _checkinTimes[key] = isNaN(parsed) ? new Date() : parsed;
        }
        if (!_breakStates[key]) {
            _breakStates[key] = { onBreak: false, breakStart: null, totalBreakMs: 0 };
        }
    });

    // Sync break states from Firestore once, then start ticking
    syncBreakStatesFromFirestore(activeLogs).then(() => {
        if (activeLogs.length > 0) {
            _dashTimerInterval = setInterval(() => tickTimers(activeLogs), 1000);
            tickTimers(activeLogs); // immediate first tick
        }
    });
}

async function syncBreakStatesFromFirestore(todayLogs) {
    for (const r of todayLogs) {
        try {
            const doc = await firestore.collection('attendance').doc(r.name).get();
            if (doc.exists) {
                const d = doc.data();
                const bs = _breakStates[r.name] || { onBreak: false, breakStart: null, totalBreakMs: 0 };
                bs.totalBreakMs = d.totalBreakMs || 0;
                bs.onBreak = d.onBreak || false;
                if (bs.onBreak && d.breakStart) {
                    bs.breakStart = d.breakStart.toDate ? d.breakStart.toDate() : new Date(d.breakStart);
                } else {
                    bs.breakStart = null;
                }
                _breakStates[r.name] = bs;
                updateBreakBtn(r.name, bs.onBreak);
            }
        } catch (e) { /* ignore */ }
    }
}

function tickTimers(todayLogs) {
    const now = new Date();
    todayLogs.forEach(r => {
        const key = r.name;
        const domKey = key.replace(/\s+/g,'_');
        const el = document.querySelector(`#timer-${domKey} .timer-display`);
        if (!el) return;

        const checkin = _checkinTimes[key] || now;
        let elapsedMs = now - checkin;

        // Subtract total past break time
        const bs = _breakStates[key] || { onBreak: false, breakStart: null, totalBreakMs: 0 };
        let breakMs = bs.totalBreakMs || 0;
        // If currently on break, add current break segment
        if (bs.onBreak && bs.breakStart) {
            breakMs += now - bs.breakStart;
        }
        elapsedMs = Math.max(0, elapsedMs - breakMs);

        el.textContent = msToHMS(elapsedMs);
    });
}

function msToHMS(ms) {
    const totalSec = Math.floor(ms / 1000);
    const h = Math.floor(totalSec / 3600);
    const m = Math.floor((totalSec % 3600) / 60);
    const s = totalSec % 60;
    return [h, m, s].map(v => String(v).padStart(2, '0')).join(':');
}

function updateBreakBtn(name, onBreak) {
    const key = name.replace(/\s+/g,'_');
    const btn = document.getElementById('break-btn-' + key);
    const badge = document.getElementById('status-badge-' + key);
    if (!btn) return;
    if (onBreak) {
        btn.innerHTML = '<i class="ti ti-player-play"></i> Resume';
        btn.classList.add('on-break');
        if (badge) { badge.textContent = 'On Break'; badge.className = 'badge badge-amber'; }
    } else {
        btn.innerHTML = '<i class="ti ti-coffee"></i> Break';
        btn.classList.remove('on-break');
        if (badge) { badge.textContent = 'Present'; badge.className = 'badge badge-green'; }
    }
}

async function toggleBreak(name) {
    const bs = _breakStates[name] || { onBreak: false, breakStart: null, totalBreakMs: 0 };
    const now = new Date();

    if (!bs.onBreak) {
        // Start break
        bs.onBreak = true;
        bs.breakStart = now;
        _breakStates[name] = bs;
        updateBreakBtn(name, true);
        try {
            await firestore.collection('attendance').doc(name).set({
                onBreak: true,
                breakStart: firebase.firestore.FieldValue.serverTimestamp(),
                totalBreakMs: bs.totalBreakMs
            }, { merge: true });
        } catch(e) { console.error('Break start error:', e); }
    } else {
        // End break — accumulate break duration
        const breakDuration = bs.breakStart ? (now - bs.breakStart) : 0;
        bs.totalBreakMs = (bs.totalBreakMs || 0) + breakDuration;
        bs.onBreak = false;
        bs.breakStart = null;
        _breakStates[name] = bs;
        updateBreakBtn(name, false);
        try {
            await firestore.collection('attendance').doc(name).set({
                onBreak: false,
                breakStart: null,
                totalBreakMs: bs.totalBreakMs
            }, { merge: true });
        } catch(e) { console.error('Break end error:', e); }
    }
}

window.toggleBreak = toggleBreak;

/* ── Init ────────────────────────────────────────────────── */
(async function init() {
    applyStoredTheme();

    await loadSettings();
    
    console.log('📊 Loading data from Firebase...');
    await renderDashboard();
    await renderEmployees();

    const logDateEl = document.getElementById('log-date');
    if (logDateEl) logDateEl.value = todayStr();
    await renderLog();

    console.log('✅ AttendQR initialized - 100% Firebase, NO IndexedDB');
})();
/* ══════════════════════════════════════════════════════
   ECONOMY TAB — Weekly Salary Calculator
   Columns: #  |  Name  |  Wage (₹/hr)  |  THW  |  Weekly Pay
══════════════════════════════════════════════════════ */

let _ecoEmployees = [];

/* Format number as Indian currency */
function fmtINR(n) {
  return '₹' + n.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

/* Load employees from Firebase and build the table */
async function renderEconomy() {
  const list = document.getElementById('eco-list');
  if (!list) return;

  list.innerHTML = '<div class="empty-state">Loading employees…</div>';
  _ecoEmployees = await getEmployees();

  document.getElementById('eco-total-employees').textContent = _ecoEmployees.length;

  if (!_ecoEmployees.length) {
    list.innerHTML = '<div class="empty-state">No employees found. Add employees in the Employees tab first.</div>';
    document.getElementById('eco-grand-total').style.display = 'none';
    updateEcoSummary();
    return;
  }

  // Fetch today's attendance to pre-fill totalHours
  const allAttendance = await getAttendance();
  const today = todayStr();
  // Build a map: empName -> totalHours for today
  const hoursMap = {};
  allAttendance.forEach(r => {
    if ((r.dateISO === today || r.date === new Date(today + 'T00:00:00').toLocaleDateString('en-IN')) && r.totalHours > 0) {
      hoursMap[r.name] = r.totalHours;
    }
  });

  list.innerHTML = _ecoEmployees.map((emp, i) => {
    const prefilled = hoursMap[emp.name] || '';
    const hasTime = prefilled !== '';
    return `
    <div class="eco-row">
      <span class="eco-serial">${i + 1}</span>
      <span>
        <div class="eco-name">${emp.name}</div>
        <div class="eco-dept">${emp.dept || '—'}</div>
      </span>
      <span>
        <input type="number" min="0" step="0.5"
          id="eco-wage-${i}"
          placeholder="e.g. 80"
          oninput="calcEcoRow(${i})">
      </span>
      <span>
        <input type="number" min="0" max="168" step="0.01"
          id="eco-thw-${i}"
          value="${prefilled ? prefilled.toFixed(2) : ''}"
          placeholder="e.g. 40"
          oninput="calcEcoRow(${i})"
          style="${hasTime ? 'border-color:var(--green);background:var(--green-bg,#f0fdf4)' : ''}">
        ${hasTime ? `<div style="font-size:11px;color:var(--green);margin-top:2px">⏱ ${allAttendance.find(r=>r.name===emp.name&&r.totalTime)?.totalTime||''} today</div>` : ''}
      </span>
      <span class="eco-payout" id="eco-pay-${i}">₹0</span>
    </div>
  `}).join('');

  document.getElementById('eco-grand-total').style.display = 'grid';
  updateEcoSummary();
}

/* Recalculate one row and refresh summary */
function calcEcoRow(i) {
  const wage = parseFloat(document.getElementById(`eco-wage-${i}`)?.value) || 0;
  const thw  = parseFloat(document.getElementById(`eco-thw-${i}`)?.value)  || 0;
  const pay  = wage * thw;
  document.getElementById(`eco-pay-${i}`).textContent = fmtINR(pay);
  updateEcoSummary();
}

/* Recompute totals and grand-total row */
function updateEcoSummary() {
  let totalHrs  = 0;
  let totalPay  = 0;

  _ecoEmployees.forEach((_, i) => {
    const wage = parseFloat(document.getElementById(`eco-wage-${i}`)?.value) || 0;
    const thw  = parseFloat(document.getElementById(`eco-thw-${i}`)?.value)  || 0;
    totalHrs += thw;
    totalPay += wage * thw;
  });

  document.getElementById('eco-total-hours').textContent  = totalHrs.toFixed(1);
  document.getElementById('eco-total-payout').textContent = fmtINR(totalPay);
  document.getElementById('eco-grand-hours').textContent  = totalHrs.toFixed(1) + ' hrs';
  document.getElementById('eco-grand-payout').textContent = fmtINR(totalPay);
}

/* Export to Excel — columns: #, Name, Dept, Wage, THW, Weekly Pay */
function exportEconomyXLSX() {
  if (!_ecoEmployees.length) { alert('No data to export.'); return; }

  const rows = _ecoEmployees.map((emp, i) => {
    const wage = parseFloat(document.getElementById(`eco-wage-${i}`)?.value) || 0;
    const thw  = parseFloat(document.getElementById(`eco-thw-${i}`)?.value)  || 0;
    return {
      'SlNo'            : i + 1,
      'Name'            : emp.name,
      'Department'      : emp.dept || '',
      'Wage (₹/hr)'     : wage,
      'THW (hrs/week)'  : thw,
      'Weekly Pay (₹)'  : wage * thw
    };
  });

  // Add grand total row
  const totalHrs = rows.reduce((s, r) => s + r['THW (hrs/week)'], 0);
  const totalPay = rows.reduce((s, r) => s + r['Weekly Pay (₹)'], 0);
  rows.push({ 'SlNo': '', 'Name': 'GRAND TOTAL', 'Department': '', 'Wage (₹/hr)': '', 'THW (hrs/week)': totalHrs, 'Weekly Pay (₹)': totalPay });

  const ws = XLSX.utils.json_to_sheet(rows);

  // Column widths
  ws['!cols'] = [{ wch: 4 }, { wch: 22 }, { wch: 16 }, { wch: 13 }, { wch: 16 }, { wch: 16 }];

  const wb = XLSX.utils.book_new();
  const weekLabel = new Date().toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' });
  XLSX.utils.book_append_sheet(wb, ws, 'Weekly Salary');
  XLSX.writeFile(wb, `weekly_salary_${weekLabel.replace(/ /g,'-')}.xlsx`);
}

window.renderEconomy      = renderEconomy;
window.calcEcoRow         = calcEcoRow;
window.exportEconomyXLSX  = exportEconomyXLSX;
/* ══════════════════════════════════════════════════════
   BILLING TAB
══════════════════════════════════════════════════════ */

let _billingRows = [];
let _billingRowId = 0;

function addBillingRow(custom) {
  const select = document.getElementById('billing-item-select');
  let itemName = '';

  if (custom) {
    itemName = '';  // blank — user types it
  } else {
    itemName = select?.value || '';
    if (!itemName) { alert('Please select an item from the catalogue first.'); return; }
    select.value = '';
  }

  const id = ++_billingRowId;
  _billingRows.push({ id, name: itemName, qty: 1, price: '' });
  renderBillingTable();
}

function renderBillingTable() {
  const tbody = document.getElementById('billing-tbody');
  const empty = document.getElementById('billing-empty');
  const totalRow = document.getElementById('billing-total-row');

  if (!_billingRows.length) {
    tbody.innerHTML = '';
    empty.style.display = 'block';
    totalRow.style.display = 'none';
    return;
  }
  empty.style.display = 'none';
  totalRow.style.display = '';

  tbody.innerHTML = _billingRows.map((row, i) => {
    const isEven = i % 2 === 0;
    return `
    <tr style="background:${isEven ? '#f8fafc' : '#fff'}">
      <td style="padding:9px 12px;color:#888;font-size:13px">${i + 1}</td>
      <td style="padding:9px 12px">
        <input
          class="billing-input billing-name"
          value="${row.name.replace(/"/g, '&quot;')}"
          placeholder="Item description…"
          style="width:100%;border:none;background:transparent;font-size:14px;color:#111;outline:none;font-family:inherit"
          oninput="updateBillingRow(${row.id}, 'name', this.value)"
        >
      </td>
      <td style="padding:9px 12px">
        <input
          type="number" min="1"
          class="billing-input"
          value="${row.qty}"
          style="width:60px;border:none;background:transparent;font-size:14px;color:#111;outline:none;text-align:right;font-family:inherit"
          oninput="updateBillingRow(${row.id}, 'qty', this.value)"
        >
      </td>
      <td style="padding:9px 12px;text-align:right">
        <input
          type="number" min="0" step="0.01"
          class="billing-input"
          value="${row.price}"
          placeholder="—"
          style="width:90px;border:none;border-bottom:1px solid #ccc;background:transparent;font-size:14px;color:#111;outline:none;text-align:right;font-family:inherit"
          oninput="updateBillingRow(${row.id}, 'price', this.value)"
        >
      </td>
      <td class="no-print" style="padding:9px 6px;text-align:center">
        <button onclick="removeBillingRow(${row.id})"
          style="background:none;border:none;cursor:pointer;color:#ef4444;font-size:16px;line-height:1;padding:2px 6px"
          title="Remove">✕</button>
      </td>
    </tr>`;
  }).join('');

  updateBillingTotal();

  // Set today's date
  const dateEl = document.getElementById('bill-date');
  if (dateEl && dateEl.textContent.trim() === '') {
    dateEl.textContent = new Date().toLocaleDateString('en-IN');
  }
}

function updateBillingRow(id, field, value) {
  const row = _billingRows.find(r => r.id === id);
  if (!row) return;
  row[field] = field === 'qty' ? (parseInt(value) || 1) : value;
  updateBillingTotal();
}

function removeBillingRow(id) {
  _billingRows = _billingRows.filter(r => r.id !== id);
  renderBillingTable();
}

function updateBillingTotal() {
  let total = 0;
  _billingRows.forEach(r => {
    const p = parseFloat(r.price) || 0;
    const q = parseInt(r.qty) || 1;
    total += p * q;
  });
  const el = document.getElementById('billing-total-val');
  if (el) el.textContent = total.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function clearBill() {
  if (_billingRows.length && !confirm('Clear all bill items?')) return;
  _billingRows = [];
  _billingRowId = 0;
  const dateEl = document.getElementById('bill-date');
  if (dateEl) dateEl.textContent = '';
  renderBillingTable();
}

window.addBillingRow = addBillingRow;
window.updateBillingRow = updateBillingRow;
window.removeBillingRow = removeBillingRow;
window.clearBill = clearBill;