let db = {};

let backendActive = false;
let backendConfig = null;

// Pagination & Sorting variables
let invCurrentPage = 1;
let invRowsPerPage = 10;
let invSortColumn = 'sno';
let invSortDirection = 'asc';

let gpSortColumn = 'out_date';
let gpSortDirection = 'desc';

// Chart instances
let makeChartInstance = null;
let statusChartInstance = null;

// ==========================================================================
// Safe Helper Functions & Clipboard
// ==========================================================================
function safeStr(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

function safeLower(val) {
    return safeStr(val).toLowerCase();
}

function safeUpper(val) {
    return safeStr(val).toUpperCase();
}

// Escape any value before inserting it into innerHTML (data comes from Sheets / Excel / user input)
function escapeHtml(val) {
    return safeStr(val)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// Local (IST) calendar date as DD/MM/YYYY.
function formatLocalDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    // SheetJS date cells can land a few seconds before local midnight (e.g. 23:59:59.999);
    // round to the nearest minute so they don't fall on the previous day.
    d = new Date(Math.round(d.getTime() / 60000) * 60000);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${day}/${m}/${y}`;
}

// Local (IST) calendar date as YYYY-MM-DD (for HTML <input type="date"> values and prefixes).
function formatIsoDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    d = new Date(Math.round(d.getTime() / 60000) * 60000);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

// Copy of a record with every text value HTML-escaped (display only, never save it back)
function escapeRecord(rec) {
    const out = {};
    Object.keys(rec || {}).forEach(k => {
        const v = rec[k];
        out[k] = (typeof v === 'string' || typeof v === 'number') ? escapeHtml(v) : v;
    });
    return out;
}

function todayLocal() {
    return formatLocalDate(new Date());
}

function todayIso() {
    return formatIsoDate(new Date());
}

// Normalise any date-ish value (Date object, Excel serial, ISO timestamp, YYYY-MM-DD, DD/MM/YYYY, Date.toString) to DD/MM/YYYY.
// Non-date text such as "Return to depot" / "NA" is kept as-is.
function normalizeDateValue(val) {
    if (val === null || val === undefined || val === '') return '';
    if (val instanceof Date) return formatLocalDate(val);
    if (typeof val === 'number') {
        if (val > 20000 && val < 80000) {
            const d = new Date(Math.round((val - 25569) * 86400) * 1000);
            if (!isNaN(d.getTime())) {
                return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
            }
        }
        return String(val);
    }
    const str = String(val).trim();
    // Full ISO timestamp (e.g. "2024-02-27T18:30:00.000Z" from Apps Script) -> local calendar DD/MM/YYYY
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(str)) {
        const d = new Date(str);
        if (!isNaN(d.getTime())) return formatLocalDate(d);
    }
    let m = str.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s.*)?$/);
    if (m) return `${m[3].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[1]}`;
    m = str.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:\s.*)?$/);
    if (m) {
        const p1 = Number(m[1]), p2 = Number(m[2]);
        if (p2 > 12 && p1 <= 12) return `${m[2].padStart(2, '0')}/${m[1].padStart(2, '0')}/${m[3]}`;
        return `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`;
    }
    // JS Date.toString() / toUTCString() output, e.g. "Wed Jul 22 2026 00:00:00 GMT+0530 (India Standard Time)"
    if (/^[A-Za-z]{3},?\s+(?:[A-Za-z]{3}\s+\d{1,2}|\d{1,2}\s+[A-Za-z]{3})\s+\d{4}/.test(str)) {
        const d = new Date(str);
        if (!isNaN(d.getTime())) return formatLocalDate(d);
    }
    return str;
}

function formatDisplayDate(val) {
    return normalizeDateValue(val);
}

function toIsoDate(val) {
    const norm = normalizeDateValue(val);
    const m = safeStr(norm).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? `${m[3]}-${m[2]}-${m[1]}` : '';
}

function safeDateTimestamp(val) {
    if (val instanceof Date) return isNaN(val.getTime()) ? 0 : new Date(val.getFullYear(), val.getMonth(), val.getDate()).getTime();
    if (!val || typeof val !== 'string') return 0;
    val = val.trim();
    if (!val || val === '—' || val === '-' || val.toLowerCase().includes('return') || val.toLowerCase() === 'nil' || val.toLowerCase() === 'na') {
        return 0;
    }
    const iso = toIsoDate(val);
    if (iso) {
        const parts = iso.split('-').map(Number);
        return new Date(parts[0], parts[1] - 1, parts[2]).getTime() || 0;
    }
    const t = new Date(val).getTime();
    return isNaN(t) ? 0 : t;
}

function copyToClipboard(text, label = 'Text') {
    if (!text) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => {
            showToast(`${label} copied to clipboard!`);
        }).catch(() => fallbackCopy(text, label));
    } else {
        fallbackCopy(text, label);
    }
}

function fallbackCopy(text, label) {
    const tempInput = document.createElement('input');
    tempInput.value = text;
    document.body.appendChild(tempInput);
    tempInput.select();
    try {
        document.execCommand('copy');
        showToast(`${label} copied to clipboard!`);
    } catch (e) {
        prompt("Copy to clipboard: Ctrl+C, Enter", text);
    }
    tempInput.remove();
}

function populateHardwareSelects() {
    // Preserved for compatibility
}

function normalizeRecord(item) {
    if (!item || typeof item !== 'object') return {};

    const getVal = (...keys) => {
        for (const k of keys) {
            if (item[k] !== undefined && item[k] !== null && item[k] !== '') return item[k];
            const lowerK = k.toLowerCase().replace(/[\s_.-]/g, '');
            for (const actualKey of Object.keys(item)) {
                if (actualKey.toLowerCase().replace(/[\s_.-]/g, '') === lowerK) {
                    if (item[actualKey] !== undefined && item[actualKey] !== null && item[actualKey] !== '') {
                        return item[actualKey];
                    }
                }
            }
        }
        return '';
    };

    return {
        ...item,
        sno: getVal('sno', 'SNo', 'SNO', 'S.No') || 0,
        type: safeStr(getVal('type', 'Type', 'category', 'Category')) || 'Pressure',
        variant: safeStr(getVal('variant', 'Variant', 'hp_type', 'HP Type')),
        make: safeStr(getVal('make', 'Make')).toUpperCase(),
        model: safeStr(getVal('model', 'Model')),
        serial_no: safeStr(getVal('serial_no', 'Serial No', 'SerialNo', 'serial_number')),
        receipt_date: normalizeDateValue(getVal('receipt_date', 'Receipt Date', 'received_date', 'Received Date')),
        received_from: safeStr(getVal('received_from', 'Received From')),
        received_date: normalizeDateValue(getVal('received_date', 'Received Date', 'receipt_date')),
        qty: getVal('qty', 'Qty', 'quantity') || 0,
        status: safeStr(getVal('status', 'Status')) || 'N',
        remarks: safeStr(getVal('remarks', 'Remarks')),
        emr_no: safeStr(getVal('emr_no', 'EMR No', 'emr_number', 'gate_pass_no', 'Gate Pass No')),
        emr_date: normalizeDateValue(getVal('emr_date', 'EMR Date', 'gate_pass_date', 'Gate Pass Date')),
        destination: safeStr(getVal('destination', 'Destination')),
        vendor: safeStr(getVal('vendor', 'Vendor')),
        out_date: normalizeDateValue(getVal('out_date', 'Out Date')),
        gate_pass_no: safeStr(getVal('gate_pass_no', 'Gate Pass No', 'emr_no', 'EMR No')),
        gate_pass_date: normalizeDateValue(getVal('gate_pass_date', 'Gate Pass Date', 'emr_date', 'EMR Date')),
        dl_date: normalizeDateValue(getVal('dl_date', 'Letter Date')),
        dispatched_on: normalizeDateValue(getVal('dispatched_on', 'Dispatched On')),
        item: safeStr(getVal('item', 'Item', 'material', 'Material')),
        tank_capacity: safeStr(getVal('tank_capacity', 'Tank Capacity', 'capacity', 'Capacity')),
        material: safeStr(getVal('material', 'Material', 'item', 'Item'))
    };
}

function formatStatusBadge(statusVal) {
    const s = safeStr(statusVal).toUpperCase();
    if (s === 'N' || s === 'NEW' || s === 'AT DEPOT' || s === 'DEPOT' || s === 'IN DEPOT' || s === 'IN STORE') return '<span class="tag tag-new">New</span>';
    if (s === 'U' || s === 'USED' || s === 'REPAIRED') return '<span class="tag tag-used">Used</span>';
    if (s === 'S' || s.includes('UNSERVICE') || s === 'SCRAP' || s === 'FAULTY') return '<span class="tag tag-unsv">Unserviceable</span>';
    if (s === 'O' || s === 'OUTWARD' || s === 'DISPATCHED' || s === 'ISSUED') return '<span class="tag tag-out">Dispatched</span>';
    return `<span class="tag tag-plain">${escapeHtml(statusVal) || 'In store'}</span>`;
}

// ==========================================================================
// Initialization & Global Event Listeners
// ==========================================================================
document.addEventListener('DOMContentLoaded', async () => {
    // Current date in DD/MM/YYYY
    const currentDateElem = document.getElementById('current-date');
    if (currentDateElem) {
        currentDateElem.textContent = todayLocal();
    }

    // Set theme text & icons
    const currentTheme = localStorage.getItem('hpcl_theme') || 'light';
    document.documentElement.setAttribute('data-theme', currentTheme);
    const themeBtnSpan = document.querySelector('#theme-toggle-btn span');
    if (themeBtnSpan) {
        themeBtnSpan.textContent = currentTheme === 'dark' ? 'Light mode' : 'Dark mode';
    }

    // Global Search Shortcut (Ctrl+K or /)
    window.addEventListener('keydown', (e) => {
        if ((e.ctrlKey && e.key.toLowerCase() === 'k') || (e.key === '/' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA')) {
            e.preventDefault();
            const searchInput = document.getElementById('global-search');
            if (searchInput) {
                searchInput.focus();
                searchInput.select();
            }
        }
    });

    const clearSearchBtn = document.getElementById('clear-search-btn');
    const globalSearchInput = document.getElementById('global-search');
    if (clearSearchBtn && globalSearchInput) {
        clearSearchBtn.addEventListener('click', () => {
            globalSearchInput.value = '';
            clearSearchBtn.style.display = 'none';
            invCurrentPage = 1;
            renderInventoryList();
            renderInwardLedger();
            renderOutwardLedger();
        });
        globalSearchInput.addEventListener('input', () => {
            clearSearchBtn.style.display = globalSearchInput.value ? 'block' : 'none';
        });
    }

    // 1. Detect Backend & Load Database
    await detectBackend();
    await loadDatabase();

    // 2. Setup Event Listeners
    setupNavigation();
    setupMobileNavigation();
    setupFilters();
    setupModals();
    setupDataActions();
    setupExcelImport();
    setupBackendControls();
    setupGatePassHandlers();
    setupTableSorting();
    
    // Set initial date values in forms (HTML <input type="date"> requires YYYY-MM-DD)
    const todayStr = todayIso();
    const formDateIds = ['inv-receipt-date', 'in-date', 'out-emr-date', 'out-date', 'gp-out-date'];
    formDateIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = todayStr;
    });
});

// ==========================================================================
// Backend Detection & State Management
// ==========================================================================
async function detectBackend() {
    let savedUrl = localStorage.getItem('hpcl_sheets_url') || 'https://docs.google.com/spreadsheets/d/1Rlyd-8qgXJwu6vaCgRBoQFTZtv9UjgmOMXG_29JJQEw/edit?gid=990657005#gid=990657005';
    if (savedUrl.includes('/macros/library/')) {
        savedUrl = 'https://docs.google.com/spreadsheets/d/1bI3FN3QprFsXRV5qx0xaW5Xk2zbJsusrJVYpYerO3BY/edit?usp=sharing';
        localStorage.setItem('hpcl_sheets_url', savedUrl);
    }
    // Remember the default link so the status shown matches the link in the box
    if (!localStorage.getItem('hpcl_sheets_url')) localStorage.setItem('hpcl_sheets_url', savedUrl);
    const savedAutoSync = localStorage.getItem('hpcl_sheets_auto_sync') !== 'false';

    const urlElem = document.getElementById('sheets-spreadsheet-id');
    const autoSyncElem = document.getElementById('sheets-auto-sync');
    if (urlElem) urlElem.value = savedUrl;
    if (autoSyncElem) autoSyncElem.checked = savedAutoSync;

    window.addEventListener('online', () => {
        showToast("Internet connected! Auto-syncing to Google Sheets...");
        updateSheetsStatusBadge();
        autoPushGoogleSheets();
    });
    window.addEventListener('offline', () => {
        showToast("Working Offline (Saved in Browser Local Storage)");
        updateSheetsStatusBadge();
    });

    try {
        const configResponse = await fetch('/api/config');
        if (configResponse.ok) {
            backendActive = true;
            const configData = await configResponse.json();
            backendConfig = configData.config;
            if (!savedUrl && backendConfig.sheet_id) {
                if (urlElem) urlElem.value = backendConfig.sheet_id;
                localStorage.setItem('hpcl_sheets_url', backendConfig.sheet_id);
            }
        }
    } catch (e) {
        backendActive = false;
    }

    updateSheetsStatusBadge();
}

async function loadDatabase() {
    try {
        const localData = localStorage.getItem('hpcl_inventory_db');
        if (localData) {
            try {
                db = JSON.parse(localData);
            } catch (parseErr) {
                console.error('Saved local database is corrupted; loading baseline instead.', parseErr);
                localStorage.setItem('hpcl_inventory_db_corrupt_backup', localData);
                localStorage.removeItem('hpcl_inventory_db');
                db = {};
            }
        } else if (backendActive) {
            try {
                const response = await fetch('/api/data');
                if (response.ok) {
                    db = await response.json();
                    localStorage.setItem('hpcl_inventory_db', JSON.stringify(db));
                }
            } catch (err) {
                console.warn("Failed loading from backend API. Loading baseline JSON.");
            }
        }

        if (!db || !db.inventory || (!db.inventory.pressure_dus && !db.inventory.stps)) {
            await loadLocalDatabaseFallback();
            return;
        }

        standardizeDatabase();
        refreshAllViews();
    } catch (error) {
        console.error("Failed to load inventory database:", error);
    }
}

async function loadLocalDatabaseFallback() {
    const localData = localStorage.getItem('hpcl_inventory_db');
    
    if (typeof window.baselineData === 'undefined' || window.baselineData === null) {
        try {
            const response = await fetch('data.json');
            if (response.ok) {
                window.baselineData = await response.json();
            }
        } catch (err) {
            console.warn("Could not load data.json:", err);
        }
    }
    
    let parsedLocal = null;
    if (localData) {
        try { parsedLocal = JSON.parse(localData); } catch (e) { parsedLocal = null; }
    }
    if (parsedLocal && parsedLocal.inventory) {
        db = parsedLocal;
    } else if (window.baselineData) {
        db = JSON.parse(JSON.stringify(window.baselineData));
    } else {
        db = {
            inventory: { pressure_dus: [], stps: [] },
            inword: [], stp_outword: [], du_outword: [], tank_outword: [], others_outword: [], gate_passes: []
        };
    }

    standardizeDatabase();
    refreshAllViews();
}

function standardizeDatabase() {
    if (!db) db = {};
    if (!db.inventory) db.inventory = {};
    if (!Array.isArray(db.inventory.pressure_dus)) db.inventory.pressure_dus = [];
    if (!Array.isArray(db.inventory.stps)) db.inventory.stps = [];
    if (!Array.isArray(db.inword)) db.inword = [];
    if (!Array.isArray(db.stp_outword)) db.stp_outword = [];
    if (!Array.isArray(db.du_outword)) db.du_outword = [];
    if (!Array.isArray(db.tank_outword)) db.tank_outword = [];
    if (!Array.isArray(db.others_outword)) db.others_outword = [];
    if (!Array.isArray(db.hanging_hardware)) db.hanging_hardware = [];
    if (!Array.isArray(db.gate_passes)) db.gate_passes = [];
    if (!Array.isArray(db.dispatch_letters)) db.dispatch_letters = [];

    const DATE_FIELDS = ['receipt_date', 'received_date', 'gate_pass_date', 'emr_date', 'out_date', 'dl_date', 'dispatched_on'];
    const TEXT_FIELDS = ['serial_no', 'emr_no', 'gate_pass_no', 'vendor', 'destination', 'make', 'model',
        'vehicle_no', 'driver', 'remarks', 'variant', 'hp_type', 'category', 'item', 'material',
        'tank_capacity', 'received_from', 'status', 'type', 'dl_no', 'vehicle_no', 'issued_by', 'purpose'];

    // Records pulled from Google Sheets / Excel can carry numbers, Date objects or ISO timestamps.
    // Coerce them once here so every view can treat them as plain strings.
    const cleanRecord = (rec) => {
        if (!rec || typeof rec !== 'object') return rec;
        DATE_FIELDS.forEach(f => {
            if (rec[f] !== undefined && rec[f] !== null && rec[f] !== '') rec[f] = normalizeDateValue(rec[f]);
        });
        // Also normalize any alternate-cased date keys pulled from Sheets / older backups
        Object.keys(rec).forEach(k => {
            if ((/(?:_date|\bdate\b|dispatched_on)$/i.test(k)) && rec[k] !== undefined && rec[k] !== null && rec[k] !== '') {
                rec[k] = normalizeDateValue(rec[k]);
            }
        });
        TEXT_FIELDS.forEach(f => {
            if (rec[f] !== undefined && rec[f] !== null && typeof rec[f] !== 'string') rec[f] = String(rec[f]);
        });
        if (typeof rec.serial_no === 'string') rec.serial_no = rec.serial_no.trim();
        return rec;
    };

    const allLists = [db.inventory.pressure_dus, db.inventory.stps, db.inword, db.stp_outword,
        db.du_outword, db.tank_outword, db.others_outword, db.gate_passes, db.dispatch_letters];

    // JSON columns that come back from Google Sheets as text
    const parseJsonField = (rec, field, target) => {
        if (typeof rec[field] === 'string' && rec[field].trim()) {
            try { rec[target] = JSON.parse(rec[field]); } catch (e) { /* keep as is */ }
        }
        if (field !== target) delete rec[field];
    };
    db.gate_passes.forEach(g => {
        parseJsonField(g, 'photos_json', 'photos');
        parseJsonField(g, 'other_items_json', 'other_items');
        if (typeof g.photos === 'string') parseJsonField(g, 'photos', 'photos');
        if (!Array.isArray(g.photos)) g.photos = [];
    });
    db.dispatch_letters.forEach(l => {
        parseJsonField(l, 'items_json', 'items');
        parseJsonField(l, 'other_items_json', 'other_items');
        if (!Array.isArray(l.items)) l.items = [];
        if (!Array.isArray(l.other_items)) l.other_items = [];
        l.items.forEach(cleanRecord);
        l.status = safeLower(l.status) || 'issued';
    });
    allLists.forEach(list => list.forEach(cleanRecord));

    // Remove fully blank rows that can come back from a sheet
    db.inventory.pressure_dus = db.inventory.pressure_dus.filter(r => r && safeStr(r.serial_no));
    db.inventory.stps = db.inventory.stps.filter(r => r && safeStr(r.serial_no));

    // Guarantee every in-store asset has a unique numeric S.No (edit buttons rely on it)
    [db.inventory.pressure_dus, db.inventory.stps].forEach(list => {
        const seen = new Set();
        let maxSno = list.reduce((m, r) => Math.max(m, Number(r.sno) || 0), 0);
        list.forEach(r => {
            let n = Number(r.sno);
            if (!n || seen.has(n)) {
                maxSno += 1;
                n = maxSno;
            }
            r.sno = n;
            seen.add(n);
        });
    });

    // EMR vouchers: rebuild the multi-asset "items" array if it was flattened by a sheet round-trip
    db.gate_passes.forEach(g => {
        if (typeof g.items === 'string') {
            try { g.items = JSON.parse(g.items); } catch (e) { g.items = null; }
        }
        if ((!Array.isArray(g.items) || g.items.length === 0) && g.items_json) {
            try { g.items = JSON.parse(g.items_json); } catch (e) { g.items = null; }
        }
        if (!Array.isArray(g.items) || g.items.length === 0) {
            const emr = safeLower(g.emr_no || g.gate_pass_no);
            const serials = safeStr(g.serial_no).split(',').map(s => s.trim()).filter(Boolean);
            const rebuilt = [];
            serials.forEach(ser => {
                const du = db.du_outword.find(d => safeLower(d.serial_no) === safeLower(ser) && (!emr || safeLower(d.emr_no || d.gate_pass_no) === emr));
                const stp = du ? null : db.stp_outword.find(s => safeLower(s.serial_no) === safeLower(ser) && (!emr || safeLower(s.emr_no || s.gate_pass_no) === emr));
                if (du) {
                    rebuilt.push({ category: 'DU', serial_no: du.serial_no, make: du.make || '', model: du.model || '', variant: du.variant || 'MPD', type: du.type || 'Pressure' });
                } else if (stp) {
                    rebuilt.push({ category: 'STP', serial_no: stp.serial_no, make: stp.make || '', model: stp.model || '', variant: `${safeStr(stp.hp_type).replace(/\s*HP$/i, '') || '0.75'} HP`, hp_type: stp.hp_type || '' });
                } else if (serials.length === 1) {
                    rebuilt.push({ category: safeStr(g.category).includes('STP') ? 'STP' : 'DU', serial_no: ser, make: g.make || '', model: g.model || '', variant: g.variant || '' });
                }
            });
            g.items = rebuilt.length > 0 ? rebuilt : null;
        }
        if (Array.isArray(g.items)) g.items.forEach(cleanRecord);
        delete g.items_json;
        if (!g.items) delete g.items;
    });
}

function saveState() {
    try {
        localStorage.setItem('hpcl_inventory_db', JSON.stringify(db));
    } catch (err) {
        console.error('Local save failed:', err);
        alert('Could not save to browser storage (storage may be full). Please export a JSON backup from Backup & Sync.');
    }
    scheduleAutoPush();
    if (backendActive) {
        saveStateToBackend();
    }
    updateDatabaseDiagnostics();
    updateSidebarBadges();
}

function isAppsScriptUrl(url) {
    const u = safeStr(url);
    return u.includes('script.google.com') && (u.includes('/exec') || u.includes('/macros/s/'));
}

// Several saves in quick succession (e.g. an Excel import) are pushed once.
let autoPushTimer = null;
function scheduleAutoPush() {
    if (autoPushTimer) clearTimeout(autoPushTimer);
    autoPushTimer = setTimeout(() => {
        autoPushTimer = null;
        autoPushGoogleSheets();
    }, 1500);
}

async function autoPushGoogleSheets() {
    const sheetsUrl = localStorage.getItem('hpcl_sheets_url');
    const autoSync = localStorage.getItem('hpcl_sheets_auto_sync') !== 'false';

    if (!sheetsUrl || !autoSync) return;

    // A plain spreadsheet link can only be read; pushing needs the Apps Script /exec Web App URL.
    if (!isAppsScriptUrl(sheetsUrl)) return;

    if (!navigator.onLine) {
        updateSheetsStatusBadge();
        return;
    }

    try {
        await pushToGoogleSheets(sheetsUrl, db);
    } catch (err) {
        console.warn("Auto-sync to Google Sheets failed:", err);
    }
}

async function pullFromGoogleSheets(url) {
    if (!url) {
        throw new Error("Please enter your Google Sheet URL or Web App Link first.");
    }
    url = url.trim();

    if (url.includes('/macros/library/')) {
        throw new Error("Invalid Link: You pasted an Apps Script Library link (/macros/library/).\n\nPlease paste either:\n1. Your Google Sheet URL (e.g. https://docs.google.com/spreadsheets/d/...)\nOR\n2. Your deployed Web App URL (ending in /exec) from Google Apps Script.");
    }

    if (typeof backendActive !== 'undefined' && backendActive) {
        try {
            const response = await fetch('/api/pull-sheets', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sheet_id: url })
            });
            if (response.ok) {
                const result = await response.json();
                if (result && result.data) return result.data;
            }
        } catch (e) {
            console.warn("Backend pull failed, falling back to direct browser fetch:", e);
        }
    }

    let spreadsheetId = null;
    if (url.includes('docs.google.com/spreadsheets')) {
        const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) spreadsheetId = match[1];
    } else if (!url.includes('http') && url.length > 20 && !url.includes('/')) {
        spreadsheetId = url;
    }

    if (spreadsheetId) {
        const dataObj = await pullFromSpreadsheetCsvEndpoints(spreadsheetId);
        if (dataObj) return dataObj;
    }

    if (url.includes('script.google.com')) {
        if (!url.includes('/exec') && !url.includes('/macros/s/')) {
            throw new Error("Invalid Apps Script URL: Please click 'Deploy' -> 'New Deployment' -> 'Web App' in Google Apps Script and copy the Web App URL (ending in /exec).");
        }

        let pullUrl = url;
        if (!pullUrl.includes('?')) {
            pullUrl += '?action=pull';
        } else if (!pullUrl.includes('action=pull')) {
            pullUrl += '&action=pull';
        }

        try {
            const response = await fetch(pullUrl, {
                method: 'GET',
                redirect: 'follow'
            });
            if (response.ok) {
                const text = await response.text();
                let parsed = null;
                try { parsed = JSON.parse(text); } catch (e) {}
                if (parsed) {
                    const dataObj = parsed.data || parsed;
                    if (dataObj.inventory || dataObj.inword || dataObj.pressure_dus || dataObj.stps) {
                        return dataObj;
                    }
                }
            }
        } catch (err) {
            console.warn("GET request to Apps Script failed, trying POST request...", err);
        }

        try {
            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ action: 'pull' })
            });
            if (response.ok) {
                const text = await response.text();
                let parsed = JSON.parse(text);
                const dataObj = parsed.data || parsed;
                if (dataObj.inventory || dataObj.inword || dataObj.pressure_dus || dataObj.stps) {
                    return dataObj;
                }
            }
        } catch (err) {
            console.warn("POST pull to Apps Script failed:", err);
        }
    }

    throw new Error("Unable to pull data from the provided URL.\n\nQuick Fix:\n1. If using Google Sheet URL: Set link sharing to 'Anyone with the link can view'.\n2. If using Apps Script: Deploy as Web App with access set to 'Anyone', and copy the /exec link.");
}

async function pullFromSpreadsheetCsvEndpoints(spreadsheetId) {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
    try {
        const response = await fetch(exportUrl);
        if (response.ok) {
            const buffer = await response.arrayBuffer();
            if (buffer && buffer.byteLength > 0 && typeof XLSX !== 'undefined') {
                const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
                const parsedDb = parseWorkbookToDb(workbook);
                if (parsedDb) {
                    return parsedDb;
                }
            }
        }
    } catch (e) {
        console.warn("Direct XLSX export pull failed, attempting CSV tabs:", e);
    }

    const sheetsToFetch = [
        { names: ["Pressure_DUs", "Inventory", "Pressure DUs", "DUs"], target: ["inventory", "pressure_dus"] },
        { names: ["STPs", "STP"], target: ["inventory", "stps"] },
        { names: ["Inward", "Inword"], target: ["inword"] },
        { names: ["STP_Outward", "STP_Outword", "STP OUTWORD"], target: ["stp_outword"] },
        { names: ["DU_Outward", "DU_Outword", "DU OUTWORD"], target: ["du_outword"] },
        { names: ["Tank_Outward", "Tank_Outword", "TANK OUTWORD"], target: ["tank_outword"] },
        { names: ["Others_Outward", "Others_Outword", "OTHERS OUTWORD"], target: ["others_outword"] },
        { names: ["EMR", "EMRs", "EMR_Records", "Gate_Passes", "Gate Passes"], target: ["gate_passes"] },
        { names: ["Dispatch_Letters", "Dispatch Letters", "Letters"], target: ["dispatch_letters"] }
    ];

    let pulledDb = {
        inventory: { pressure_dus: [], stps: [] },
        inword: [], stp_outword: [], du_outword: [], tank_outword: [], others_outword: [], gate_passes: [], dispatch_letters: [], hanging_hardware: []
    };

    let fetchedCount = 0;

    for (const sheetDef of sheetsToFetch) {
        for (const sheetName of sheetDef.names) {
            const csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
            try {
                const res = await fetch(csvUrl);
                if (res.ok) {
                    const csvText = await res.text();
                    if (csvText && csvText.trim().length > 0 && typeof XLSX !== 'undefined') {
                        const workbook = XLSX.read(csvText, { type: 'string', raw: true });
                        const firstSheetName = workbook.SheetNames[0];
                        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
                        if (rows && rows.length > 0) {
                            if (sheetDef.target.length === 2) {
                                pulledDb[sheetDef.target[0]][sheetDef.target[1]] = rows;
                            } else {
                                pulledDb[sheetDef.target[0]] = rows;
                            }
                            fetchedCount++;
                            break;
                        }
                    }
                }
            } catch (e) {
                console.warn(`Could not pull tab ${sheetName}:`, e);
            }
        }
    }

    return fetchedCount > 0 ? pulledDb : null;
}

function parseWorkbookToDb(workbook) {
    let newDb = {
        inventory: { pressure_dus: [], stps: [] },
        inword: [], stp_outword: [], du_outword: [], tank_outword: [], others_outword: [], gate_passes: [], hanging_hardware: []
    };

    let foundAny = false;

    function getSheetRows(possibleNames) {
        for (const name of possibleNames) {
            for (const sname of workbook.SheetNames) {
                if (sname.toLowerCase().trim() === name.toLowerCase().trim()) {
                    return XLSX.utils.sheet_to_json(workbook.Sheets[sname]);
                }
            }
        }
        return [];
    }

    const pdus = getSheetRows(["Pressure_DUs", "Inventory", "Pressure DUs", "DUs"]);
    if (pdus.length > 0) {
        newDb.inventory.pressure_dus = pdus;
        foundAny = true;
    }

    const stps = getSheetRows(["STPs", "STP"]);
    if (stps.length > 0) {
        newDb.inventory.stps = stps;
        foundAny = true;
    }

    const inwords = getSheetRows(["Inward", "Inword"]);
    if (inwords.length > 0) {
        newDb.inword = inwords;
        foundAny = true;
    }

    const stpOut = getSheetRows(["STP_Outward", "STP_Outword", "STP OUTWORD"]);
    if (stpOut.length > 0) {
        newDb.stp_outword = stpOut;
        foundAny = true;
    }

    const duOut = getSheetRows(["DU_Outward", "DU_Outword", "DU OUTWORD"]);
    if (duOut.length > 0) {
        newDb.du_outword = duOut;
        foundAny = true;
    }

    const tankOut = getSheetRows(["Tank_Outward", "Tank_Outword", "TANK OUTWORD"]);
    if (tankOut.length > 0) {
        newDb.tank_outword = tankOut;
        foundAny = true;
    }

    const othOut = getSheetRows(["Others_Outward", "Others_Outword", "OTHERS OUTWORD"]);
    if (othOut.length > 0) {
        newDb.others_outword = othOut;
        foundAny = true;
    }

    const gps = getSheetRows(["EMR", "EMRs", "EMR_Records", "Gate_Passes", "Gate Passes"]);
    if (gps.length > 0) {
        newDb.gate_passes = gps;
        foundAny = true;
    }
    newDb.dispatch_letters = getSheetRows(["Dispatch_Letters", "Dispatch Letters", "Letters"]);

    return foundAny ? newDb : null;
}

async function pushToGoogleSheets(url, dataPayload) {
    if (!url) {
        throw new Error("Please enter an Apps Script Web App URL (ending in /exec) first.");
    }
    url = url.trim();

    if (url.includes('/macros/library/')) {
        throw new Error("You pasted an Apps Script Library link (/macros/library/). Deploy as a Web App (click 'Setup steps') and paste the /exec URL instead.");
    }

    if (url.includes('docs.google.com/spreadsheets')) {
        throw new Error("A Google Sheet link can only pull (read) data.\n\nTo push data and upload photos:\n1. Open your Google Sheet -> Extensions -> Apps Script.\n2. Click 'Copy Apps Script' here and paste it into Apps Script.\n3. Click Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).\n4. Paste the /exec Web App URL above and click 'Push now'.");
    }

    if (!isAppsScriptUrl(url)) {
        throw new Error("To push to Google Sheets, paste a deployed Apps Script Web App URL (e.g. https://script.google.com/macros/s/.../exec).\n\nClick 'Setup steps' or 'Copy Apps Script' to set it up.");
    }

    updateSheetsStatusBadge(true, "Syncing...");

    try {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(dataPayload)
        });
        const text = await res.text();
        let parsed = null;
        try { parsed = JSON.parse(text); } catch (e) { parsed = null; }
        if (!res.ok || !parsed || parsed.status === 'error') {
            const msg = (parsed && parsed.message) ? parsed.message : `Unexpected response from Apps Script (HTTP ${res.status}). Check that Web App is deployed with 'Who has access: Anyone' and 'google_apps_script.js' v3.0.`;
            throw new Error(msg);
        }
        showToast("Synced to Google Sheets!");
        updateSheetsStatusBadge();
        return parsed;
    } catch (err) {
        console.error("Error pushing to Google Sheets:", err);
        updateSheetsStatusBadge();
        if (err && err.name === 'TypeError') {
            throw new Error("Could not reach the Apps Script Web App (network/CORS block). Ensure you deployed it as a Web App with 'Execute as: Me' and 'Who has access: Anyone', and pasted the /exec URL.");
        }
        throw err;
    }
}

async function saveStateToBackend() {
    try {
        await fetch('/api/data', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(db)
        });
    } catch (error) {
        console.error("Error saving state to backend API:", error);
    }
}

// ==========================================================================
// Views & Navigation Controller
// ==========================================================================
function refreshAllViews() {
    populateFilterDropdowns();
    updateDashboardMetrics();
    renderInventoryList();
    renderInwardLedger();
    renderOutwardLedger();
    renderGatePassView();
    updateCharts();
    updateSidebarBadges();
    updateQuickChips();
    updateDatabaseDiagnostics();
}

function updateSidebarBadges() {
    const totalDUs = (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).length;
    const totalSTPs = (db.inventory && db.inventory.stps ? db.inventory.stps : []).length;
    const activeInventoryCount = totalDUs + totalSTPs;
    const inwardCount = (db.inword || []).length;
    const duOutCount = (db.du_outword || []).length;
    const stpOutCount = (db.stp_outword || []).length;
    const tankOutCount = (db.tank_outword || []).length;
    const othersOutCount = (db.others_outword || []).length;
    const totalOutwardCount = duOutCount + stpOutCount + tankOutCount + othersOutCount;
    const gatePassCount = (db.gate_passes || []).length;

    const bInv = document.getElementById('nav-badge-inventory');
    const bInw = document.getElementById('nav-badge-inward');
    const bOut = document.getElementById('nav-badge-outward');
    const bGp = document.getElementById('nav-badge-emr') || document.getElementById('nav-badge-gatepass');

    if (bInv) bInv.textContent = activeInventoryCount;
    if (bInw) bInw.textContent = inwardCount;
    if (bOut) bOut.textContent = totalOutwardCount;
    if (bGp) bGp.textContent = gatePassCount;

    // Mobile bottom navigation badges
    const mInv = document.getElementById('mobile-badge-inventory');
    const mInw = document.getElementById('mobile-badge-inward');
    const mOut = document.getElementById('mobile-badge-outward');
    const mGp = document.getElementById('mobile-badge-emr');

    if (mInv) mInv.textContent = activeInventoryCount;
    if (mInw) mInw.textContent = inwardCount;
    if (mOut) mOut.textContent = totalOutwardCount;
    if (mGp) mGp.textContent = gatePassCount;

    // Sub-tab badges in outward ledger
    const subDu = document.getElementById('sub-count-du');
    const subStp = document.getElementById('sub-count-stp');
    const subTank = document.getElementById('sub-count-tank');
    const subOth = document.getElementById('sub-count-others');

    if (subDu) subDu.textContent = duOutCount;
    if (subStp) subStp.textContent = stpOutCount;
    if (subTank) subTank.textContent = tankOutCount;
    if (subOth) subOth.textContent = othersOutCount;

    const waitingLetters = (db.dispatch_letters || []).filter(l => l.status === 'issued').length;
    const bLet = document.getElementById('nav-badge-letters');
    if (bLet) {
        bLet.textContent = waitingLetters;
        bLet.classList.toggle('attention', waitingLetters > 0);
        bLet.title = `${waitingLetters} letter(s) waiting at depot`;
    }
    const mLet = document.getElementById('mobile-badge-letters');
    if (mLet) {
        mLet.textContent = waitingLetters || '';
        mLet.classList.toggle('attention', waitingLetters > 0);
    }
    if (mGp) mGp.textContent = '';
    if (mInv) mInv.textContent = '';
}

function updateQuickChips() {
    const pdus = (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).map(normalizeRecord);
    const stps = (db.inventory && db.inventory.stps ? db.inventory.stps : []).map(normalizeRecord);

    const totalAll = pdus.length + stps.length;
    const countDu = pdus.length;
    const countStp = stps.length;

    const allItems = [...pdus, ...stps];
    // "New DUs" / "Used DUs" chips filter DUs only, so count DUs only
    const countNew = pdus.filter(i => {
        const s = safeUpper(i.status);
        return s === 'N' || s.includes('NEW') || s.includes('DEPOT');
    }).length;
    const countUsed = pdus.filter(i => {
        const s = safeUpper(i.status);
        return s === 'U' || s.includes('USED');
    }).length;
    const countScrap = allItems.filter(i => {
        const s = safeUpper(i.status);
        return s === 'S' || s.includes('UNSERVICEABLE') || s.includes('SCRAP');
    }).length;

    const cAll = document.getElementById('chip-count-all');
    const cDu = document.getElementById('chip-count-du');
    const cStp = document.getElementById('chip-count-stp');
    const cNew = document.getElementById('chip-count-new');
    const cUsed = document.getElementById('chip-count-used');
    const cScrap = document.getElementById('chip-count-scrap');

    if (cAll) cAll.textContent = totalAll;
    if (cDu) cDu.textContent = countDu;
    if (cStp) cStp.textContent = countStp;
    if (cNew) cNew.textContent = countNew;
    if (cUsed) cUsed.textContent = countUsed;
    if (cScrap) cScrap.textContent = countScrap;
}

function updateDatabaseDiagnostics() {
    const totalDUs = (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).length;
    const totalSTPs = (db.inventory && db.inventory.stps ? db.inventory.stps : []).length;
    const inwardCount = (db.inword || []).length;
    const outwardCount = (db.du_outword || []).length + (db.stp_outword || []).length + (db.tank_outword || []).length + (db.others_outword || []).length;
    const gpCount = (db.gate_passes || []).length;

    const rawStr = localStorage.getItem('hpcl_inventory_db') || '';
    const kbUsed = (rawStr.length / 1024).toFixed(1);

    const elDus = document.getElementById('diag-dus');
    const elStps = document.getElementById('diag-stps');
    const elInw = document.getElementById('diag-inward');
    const elOut = document.getElementById('diag-outward');
    const elGps = document.getElementById('diag-gps');
    const elStore = document.getElementById('diag-storage');

    if (elDus) elDus.textContent = totalDUs;
    if (elStps) elStps.textContent = totalSTPs;
    if (elInw) elInw.textContent = inwardCount;
    if (elOut) elOut.textContent = outwardCount;
    if (elGps) elGps.textContent = gpCount;
    if (elStore) elStore.textContent = `${kbUsed} KB used`;
    const elLet = document.getElementById('diag-letters');
    if (elLet) elLet.textContent = (db.dispatch_letters || []).length;
    const elPh = document.getElementById('diag-photos');
    if (elPh) elPh.textContent = (db.gate_passes || []).reduce((n, g) => n + (Array.isArray(g.photos) ? g.photos.length : 0), 0);
}

function setupMobileNavigation() {
    const mobileToggle = document.getElementById('mobile-menu-toggle');
    const mobileClose = document.getElementById('mobile-sidebar-close');
    const backdrop = document.getElementById('sidebar-backdrop');
    const sidebar = document.querySelector('.sidebar');

    function openDrawer() {
        if (sidebar) sidebar.classList.add('mobile-open');
        if (backdrop) backdrop.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    function closeDrawer() {
        if (sidebar) sidebar.classList.remove('mobile-open');
        if (backdrop) backdrop.classList.remove('active');
        document.body.style.overflow = '';
    }

    window.closeMobileSidebar = closeDrawer;

    if (mobileToggle) mobileToggle.addEventListener('click', openDrawer);
    if (mobileClose) mobileClose.addEventListener('click', closeDrawer);
    if (backdrop) backdrop.addEventListener('click', closeDrawer);

    // Close on Escape key
    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && sidebar && sidebar.classList.contains('mobile-open')) {
            closeDrawer();
        }
    });

    // Ensure all tables have mobile horizontal scroll hints
    document.querySelectorAll('.table-responsive').forEach(wrap => {
        if (!wrap.previousElementSibling || !wrap.previousElementSibling.classList.contains('mobile-table-hint')) {
            const hint = document.createElement('div');
            hint.className = 'mobile-table-hint';
            hint.innerHTML = '<i class="fa-solid fa-arrows-left-right"></i> Scroll table horizontally to see all columns';
            wrap.parentNode.insertBefore(hint, wrap);
        }
    });
}

function setupNavigation() {
    const navItems = document.querySelectorAll('.nav-item[data-tab], .mobile-nav-item[data-tab]');
    document.querySelectorAll('[data-open-drawer]').forEach(btn => btn.addEventListener('click', () => {
        const toggle = document.getElementById('mobile-menu-toggle');
        if (toggle) toggle.click();
    }));
    const tabPanels = document.querySelectorAll('.tab-panel');

    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const tabId = item.getAttribute('data-tab');
            
            navItems.forEach(nav => {
                if (nav.getAttribute('data-tab') === tabId) {
                    nav.classList.add('active');
                } else {
                    nav.classList.remove('active');
                }
            });

            tabPanels.forEach(panel => {
                if (panel.id === `tab-${tabId}` || (tabId === 'emr' && panel.id === 'tab-gate-pass') || (tabId === 'gate-pass' && panel.id === 'tab-emr')) {
                    panel.classList.add('active');
                } else {
                    panel.classList.remove('active');
                }
            });

            // Auto-close mobile sidebar drawer when any tab item is clicked
            if (typeof window.closeMobileSidebar === 'function') {
                window.closeMobileSidebar();
            }

            const mainContent = document.querySelector('.main-content');
            if (mainContent) mainContent.scrollTop = 0;
            window.scrollTo({ top: 0, behavior: 'smooth' });
            
            if (tabId === 'dashboard') {
                updateDashboardMetrics();
                updateCharts();
            } else if (tabId === 'inventory') {
                renderInventoryList();
            } else if (tabId === 'inward') {
                renderInwardLedger();
            } else if (tabId === 'outward') {
                renderOutwardLedger();
            } else if (tabId === 'gate-pass' || tabId === 'emr') {
                renderGatePassView();
            } else if (tabId === 'data-mgmt') {
                updateDatabaseDiagnostics();
            } else if (tabId === 'letters' && typeof renderLetters === 'function') {
                renderLetters();
            }
            try { history.replaceState(null, '', '#' + tabId); } catch (err) {}
        });
    });

    const subTabItems = document.querySelectorAll('.sub-tab-item');
    const subTabPanels = document.querySelectorAll('.sub-panel');

    subTabItems.forEach(item => {
        item.addEventListener('click', () => {
            const subtabId = item.getAttribute('data-subtab');
            
            subTabItems.forEach(tab => tab.classList.remove('active'));
            item.classList.add('active');

            subTabPanels.forEach(panel => {
                if (panel.id === `subtab-${subtabId}`) {
                    panel.classList.add('active');
                } else {
                    panel.classList.remove('active');
                }
            });
        });
    });

    const viewAllLedgersBtn = document.getElementById('view-all-ledgers');
    if (viewAllLedgersBtn) {
        viewAllLedgersBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const inwardNav = document.querySelector('.nav-item[data-tab="inward"]');
            if (inwardNav) inwardNav.click();
        });
    }
}

// ==========================================================================
// Dashboard View & Charts
// ==========================================================================
function updateDashboardMetrics() {
    const pdus = (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).map(normalizeRecord);
    const stps = (db.inventory && db.inventory.stps ? db.inventory.stps : []).map(normalizeRecord);

    const totalDUs = pdus.length;
    const newDUs = pdus.filter(d => {
        const s = safeUpper(d.status);
        return s === 'N' || s.includes('NEW') || s.includes('DEPOT');
    }).length;
    const usedDUs = pdus.filter(d => {
        const s = safeUpper(d.status);
        return s === 'U' || s.includes('USED');
    }).length;
    
    const elemTotalDus = document.getElementById('stat-total-dus');
    const elemNewDus = document.getElementById('stat-new-dus');
    const elemUsedDus = document.getElementById('stat-used-dus');
    if (elemTotalDus) elemTotalDus.textContent = totalDUs;
    if (elemNewDus) elemNewDus.textContent = newDUs;
    if (elemUsedDus) elemUsedDus.textContent = usedDUs;

    const totalSTPs = stps.length;
    const hp075 = stps.filter(s => safeStr(s.hp_type || s.variant).includes('0.75')).length;
    const hp15 = stps.filter(s => safeStr(s.hp_type || s.variant).includes('1.5')).length;
    
    const elemTotalStps = document.getElementById('stat-total-stps');
    const elemHp075 = document.getElementById('stat-075hp-stps');
    const elemHp15 = document.getElementById('stat-15hp-stps');
    if (elemTotalStps) elemTotalStps.textContent = totalSTPs;
    if (elemHp075) elemHp075.textContent = hp075;
    if (elemHp15) elemHp15.textContent = hp15;

    // Total Active Depot Stock
    const totalActiveStock = totalDUs + totalSTPs;
    const elemActiveStock = document.getElementById('stat-total-active');
    if (elemActiveStock) elemActiveStock.textContent = totalActiveStock;

    // Outward transactions count
    const stpOut = (db.stp_outword || []).length;
    const duOut = (db.du_outword || []).length;
    const tankOut = (db.tank_outword || []).length;
    const othOut = (db.others_outword || []).length;
    const totalOutpasses = stpOut + duOut + tankOut + othOut;

    const elemOut = document.getElementById('stat-total-outpasses');
    if (elemOut) elemOut.textContent = totalOutpasses;

    // Combined Recent Activity rows across all ledger streams
    const activity = [];
    
    (db.inword || []).map(normalizeRecord).forEach(item => {
        activity.push({
            date: item.received_date || item.receipt_date || '—',
            type: 'Inward',
            badgeClass: 'tag-new',
            material: item.item || 'Incoming Material',
            serialMake: item.make || '—',
            destSource: item.received_from || 'Vendor',
            qty: `${item.qty || 1} units`
        });
    });

    (db.du_outword || []).map(normalizeRecord).forEach(item => {
        activity.push({
            date: item.out_date || item.emr_date || '—',
            type: 'Out · DU',
            badgeClass: 'tag-out',
            material: item.variant || item.type || 'Dispenser DU',
            serialMake: item.serial_no ? `${item.serial_no} (${item.make || ''})` : (item.make || '—'),
            destSource: item.destination || item.vendor || '—',
            qty: '1 DU'
        });
    });

    (db.stp_outword || []).map(normalizeRecord).forEach(item => {
        activity.push({
            date: item.out_date || item.emr_date || '—',
            type: 'Out · STP',
            badgeClass: 'tag-out',
            material: `STP Pump ${item.hp_type || item.variant ? '(' + (item.hp_type || item.variant) + ' HP)' : ''}`,
            serialMake: item.serial_no ? `${item.serial_no} (${item.make || ''})` : (item.make || '—'),
            destSource: item.destination || item.vendor || '—',
            qty: '1 STP'
        });
    });

    (db.tank_outword || []).map(normalizeRecord).forEach(item => {
        activity.push({
            date: item.emr_date || item.out_date || '—',
            type: 'Out · tank',
            badgeClass: 'tag-plain',
            material: item.tank_capacity ? `Tank (${item.tank_capacity})` : 'Storage Tank',
            serialMake: item.emr_no ? `EMR: ${item.emr_no}` : '—',
            destSource: item.destination || item.vendor || '—',
            qty: '1 Tank'
        });
    });

    (db.others_outword || []).map(normalizeRecord).forEach(item => {
        activity.push({
            date: item.emr_date || item.out_date || '—',
            type: 'Out · other',
            badgeClass: 'tag-plain',
            material: item.material || item.item || 'Other Materials',
            serialMake: item.emr_no ? `EMR: ${item.emr_no}` : '—',
            destSource: item.destination || item.vendor || '—',
            qty: `${item.qty || 1} units`
        });
    });

    activity.sort((a, b) => safeDateTimestamp(b.date) - safeDateTimestamp(a.date));

    const activityBody = document.getElementById('dashboard-activity-table');
    if (activityBody) {
        activityBody.innerHTML = '';
        const topRecent = activity.slice(0, 6);
        if (topRecent.length === 0) {
            activityBody.innerHTML = `<tr><td colspan="6" class="empty-row"><strong>No movement yet</strong>Inward receipts and dispatches will appear here.</td></tr>`;
        } else {
            topRecent.map(escapeRecord).forEach(act => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="white-space:nowrap;">${formatDisplayDate(act.date) || '—'}</td>
                    <td><span class="tag ${act.badgeClass}">${act.type}</span></td>
                    <td>${act.material}</td>
                    <td><span class="serial">${act.serialMake}</span></td>
                    <td>${act.destSource}</td>
                    <td class="text-right">${act.qty}</td>
                `;
                activityBody.appendChild(tr);
            });
        }
    }
}

function updateCharts() {
    if (typeof Chart === 'undefined') return;

    const makeCounts = {
        GVR: 0,
        TOKHEIM: 0,
        FLORITE: 0,
        OTHERS: 0
    };

    (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).forEach(item => {
        const make = safeUpper(item.make);
        if (make.includes('GVR') || make.includes('GILBARCO')) makeCounts.GVR++;
        else if (make.includes('TOKHEIM')) makeCounts.TOKHEIM++;
        else makeCounts.OTHERS++;
    });

    (db.inventory && db.inventory.stps ? db.inventory.stps : []).forEach(item => {
        const make = safeUpper(item.make);
        if (make.includes('GVR')) makeCounts.GVR++;
        else if (make.includes('FLORITE')) makeCounts.FLORITE++;
        else makeCounts.OTHERS++;
    });

    const css = getComputedStyle(document.documentElement);
    const tok = (name, fallback) => (css.getPropertyValue(name) || '').trim() || fallback;
    const inkColor = tok('--ink-2', '#3D5470');
    const ruleColor = tok('--rule', '#D5DDE6');
    const surfaceColor = tok('--surface', '#ffffff');
    const fontFamily = "'Barlow', 'Segoe UI', sans-serif";
    Chart.defaults.font.family = fontFamily;
    Chart.defaults.color = inkColor;

    const makeCanvas = document.getElementById('chart-stock-makes');
    if (makeCanvas) {
        if (makeChartInstance) makeChartInstance.destroy();
        const rows = [
            ['Gilbarco (GVR)', makeCounts.GVR, tok('--blue', '#0A58B8')],
            ['Tokheim', makeCounts.TOKHEIM, tok('--red', '#D2232A')],
            ['Florite', makeCounts.FLORITE, tok('--green', '#1E8A57')],
            ['Others', makeCounts.OTHERS, tok('--ink-3', '#6B7F96')]
        ];
        makeChartInstance = new Chart(makeCanvas.getContext('2d'), {
            type: 'bar',
            data: {
                labels: rows.map(r => r[0]),
                datasets: [{
                    label: 'In store',
                    data: rows.map(r => r[1]),
                    backgroundColor: rows.map(r => r[2]),
                    borderRadius: 3,
                    barThickness: 26
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { color: ruleColor }, border: { display: false }, ticks: { precision: 0 } },
                    y: { grid: { display: false }, border: { display: false }, ticks: { font: { size: 14, weight: '600' } } }
                }
            }
        });
    }

    const duStatusCanvas = document.getElementById('chart-du-status');
    if (duStatusCanvas) {
        if (statusChartInstance) statusChartInstance.destroy();
        const statusCounts = { N: 0, U: 0, S: 0 };
        (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).forEach(item => {
            const stat = safeUpper(item.status);
            if (statusCounts[stat] !== undefined) statusCounts[stat]++;
            else if (stat.includes('NEW')) statusCounts.N++;
            else if (stat.includes('USED') || stat.includes('REPAIRED')) statusCounts.U++;
            else if (stat.includes('UNSERVICE') || stat.includes('SCRAP') || stat.includes('FAULTY')) statusCounts.S++;
        });
        statusChartInstance = new Chart(duStatusCanvas.getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: [`New ${statusCounts.N}`, `Used ${statusCounts.U}`, `Unserviceable ${statusCounts.S}`],
                datasets: [{
                    data: [statusCounts.N, statusCounts.U, statusCounts.S],
                    backgroundColor: [tok('--green', '#1E8A57'), tok('--amber', '#B76E00'), tok('--red', '#D2232A')],
                    borderWidth: 2,
                    borderColor: surfaceColor
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                cutout: '68%',
                plugins: {
                    legend: { position: 'right', labels: { boxWidth: 10, boxHeight: 10, padding: 14, font: { size: 14 } } }
                }
            }
        });
    }
}

// ==========================================================================
// Filter Dropdowns, Live Search & Sorting Setup
// ==========================================================================
function populateFilterDropdowns() {
    const makeSelect = document.getElementById('inv-make-filter');
    if (!makeSelect) return;

    const currentVal = makeSelect.value;
    const makesSet = new Set();
    
    (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).forEach(d => {
        const m = safeUpper(d.make);
        if (m) makesSet.add(m);
    });
    (db.inventory && db.inventory.stps ? db.inventory.stps : []).forEach(s => {
        const m = safeUpper(s.make);
        if (m) makesSet.add(m);
    });

    makeSelect.innerHTML = '<option value="all">All makes</option>';
    Array.from(makesSet).sort().forEach(make => {
        const opt = document.createElement('option');
        opt.value = make;
        opt.textContent = make;
        if (make === currentVal) opt.selected = true;
        makeSelect.appendChild(opt);
    });
}

function setupFilters() {
    const typeFilter = document.getElementById('inv-type-filter');
    const makeFilter = document.getElementById('inv-make-filter');
    const statusFilter = document.getElementById('inv-status-filter');
    const clearFiltersBtn = document.getElementById('clear-inv-filters');
    const globalSearch = document.getElementById('global-search');
    const inwardSearch = document.getElementById('inward-search');
    const outwardSearch = document.getElementById('outward-search');

    if (typeFilter) {
        typeFilter.addEventListener('change', () => {
            syncQuickChips();
            invCurrentPage = 1;
            renderInventoryList();
        });
    }
    if (makeFilter) {
        makeFilter.addEventListener('change', () => {
            invCurrentPage = 1;
            renderInventoryList();
        });
    }
    if (statusFilter) {
        statusFilter.addEventListener('change', () => {
            syncQuickChips();
            invCurrentPage = 1;
            renderInventoryList();
        });
    }

    if (clearFiltersBtn) {
        clearFiltersBtn.addEventListener('click', () => {
            if (typeFilter) typeFilter.value = 'all';
            if (makeFilter) makeFilter.value = 'all';
            if (statusFilter) statusFilter.value = 'all';
            if (globalSearch) globalSearch.value = '';
            const clearBtn = document.getElementById('clear-search-btn');
            if (clearBtn) clearBtn.style.display = 'none';
            syncQuickChips();
            invCurrentPage = 1;
            renderInventoryList();
        });
    }

    // Quick filter chips click handling
    const chipsContainer = document.getElementById('inventory-quick-chips');
    if (chipsContainer) {
        chipsContainer.addEventListener('click', (e) => {
            const chip = e.target.closest('.chip-btn');
            if (!chip) return;

            chipsContainer.querySelectorAll('.chip-btn').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');

            const fType = chip.getAttribute('data-filter-type') || 'all';
            const fStatus = chip.getAttribute('data-filter-status') || 'all';

            if (typeFilter) typeFilter.value = fType;
            if (statusFilter) statusFilter.value = fStatus;

            invCurrentPage = 1;
            renderInventoryList();
        });
    }

    if (globalSearch) {
        globalSearch.addEventListener('input', () => {
            invCurrentPage = 1;
            renderInventoryList();
            renderInwardLedger();
            renderOutwardLedger();
        });
    }

    if (inwardSearch) {
        inwardSearch.addEventListener('input', renderInwardLedger);
    }
    if (outwardSearch) {
        outwardSearch.addEventListener('input', renderOutwardLedger);
    }
}

function syncQuickChips() {
    const typeFilter = document.getElementById('inv-type-filter');
    const statusFilter = document.getElementById('inv-status-filter');
    const curType = typeFilter ? typeFilter.value : 'all';
    const curStat = statusFilter ? statusFilter.value : 'all';

    const chips = document.querySelectorAll('#inventory-quick-chips .chip-btn');
    chips.forEach(chip => {
        const cType = chip.getAttribute('data-filter-type');
        const cStat = chip.getAttribute('data-filter-status');
        if (cType === curType && cStat === curStat) {
            chip.classList.add('active');
        } else {
            chip.classList.remove('active');
        }
    });
}

function setupTableSorting() {
    document.querySelectorAll('.data-table thead th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const table = th.closest('table');
            const sortKey = th.getAttribute('data-sort');
            if (!table || !sortKey) return;

            if (table.id === 'inventory-table') {
                if (invSortColumn === sortKey) {
                    invSortDirection = invSortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    invSortColumn = sortKey;
                    invSortDirection = 'asc';
                }
                updateSortHeaderIcons(table, invSortColumn, invSortDirection);
                renderInventoryList();
            } else if (table.id === 'gate-pass-table') {
                if (gpSortColumn === sortKey) {
                    gpSortDirection = gpSortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    gpSortColumn = sortKey;
                    gpSortDirection = 'asc';
                }
                updateSortHeaderIcons(table, gpSortColumn, gpSortDirection);
                renderGatePassView();
            }
        });
    });
}

function updateSortHeaderIcons(table, activeColumn, direction) {
    table.querySelectorAll('thead th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        const icon = th.querySelector('.sort-icon');
        if (icon) icon.className = 'fa-solid fa-sort sort-icon';
        if (th.getAttribute('data-sort') === activeColumn) {
            th.classList.add(direction === 'asc' ? 'sort-asc' : 'sort-desc');
            if (icon) {
                icon.className = `fa-solid fa-sort-${direction === 'asc' ? 'up' : 'down'} sort-icon`;
            }
        }
    });
}

// ==========================================================================
// In-Store Inventory List & Pagination
// ==========================================================================
// Items currently visible in the In-Store table (filters + search + sort). Shared with 'Export View'.
function getInventoryViewItems() {
    const typeElem = document.getElementById('inv-type-filter');
    const makeElem = document.getElementById('inv-make-filter');
    const statusElem = document.getElementById('inv-status-filter');
    const searchElem = document.getElementById('global-search');

    const typeFilter = typeElem ? typeElem.value : 'all';
    const makeFilter = makeElem ? makeElem.value : 'all';
    const statusFilter = statusElem ? statusElem.value : 'all';
    const searchVal = searchElem ? safeLower(searchElem.value) : '';

    let items = [];
    
    if (typeFilter === 'all' || typeFilter === 'du') {
        (db.inventory && db.inventory.pressure_dus ? db.inventory.pressure_dus : []).forEach(raw => {
            const item = normalizeRecord(raw);
            items.push({
                ...item,
                category: 'Dispensing Unit (DU)',
                display_type: item.type || 'Pressure DU'
            });
        });
    }

    if (typeFilter === 'all' || typeFilter === 'stp') {
        (db.inventory && db.inventory.stps ? db.inventory.stps : []).forEach(raw => {
            const item = normalizeRecord(raw);
            items.push({
                ...item,
                type: 'STP',
                display_type: 'Submersible Pump (STP)',
                variant: item.variant ? (item.variant.includes('HP') ? item.variant : `${item.variant} HP`) : (item.hp_type ? `${item.hp_type} HP` : 'STP'),
                category: 'Submersible Pump (STP)'
            });
        });
    }

    if (makeFilter !== 'all') {
        items = items.filter(i => safeLower(i.make) === safeLower(makeFilter));
    }

    if (statusFilter !== 'all') {
        items = items.filter(i => {
            const st = safeUpper(i.status);
            if (statusFilter === 'N') return st === 'N' || st.includes('NEW');
            if (statusFilter === 'U') return st === 'U' || st.includes('USED');
            if (statusFilter === 'S') return st === 'S' || st.includes('UNSERVICEABLE') || st.includes('SCRAP');
            return st === statusFilter;
        });
    }

    if (searchVal !== '') {
        items = items.filter(i => {
            return safeLower(i.serial_no).includes(searchVal) ||
                   safeLower(i.model).includes(searchVal) ||
                   safeLower(i.variant).includes(searchVal) ||
                   safeLower(i.make).includes(searchVal) ||
                   safeLower(i.remarks).includes(searchVal) ||
                   safeLower(i.receipt_date).includes(searchVal);
        });
    }

    // Sort items
    items.sort((a, b) => {
        let valA = a[invSortColumn];
        let valB = b[invSortColumn];

        if (invSortColumn === 'sno') {
            valA = Number(valA) || 0;
            valB = Number(valB) || 0;
            return invSortDirection === 'asc' ? valA - valB : valB - valA;
        }
        if (invSortColumn === 'receipt_date') {
            return invSortDirection === 'asc' 
                ? safeDateTimestamp(valA) - safeDateTimestamp(valB)
                : safeDateTimestamp(valB) - safeDateTimestamp(valA);
        }
        valA = safeLower(valA);
        valB = safeLower(valB);
        return invSortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
    });

    return items;
}

function renderInventoryList() {
    const tableElem = document.getElementById('inventory-table');
    const tbody = tableElem ? tableElem.querySelector('tbody') : null;
    if (!tbody) return;

    tbody.innerHTML = '';

    const items = getInventoryViewItems();

    const totalItems = items.length;
    const maxPages = Math.max(Math.ceil(totalItems / invRowsPerPage), 1);
    if (invCurrentPage > maxPages) invCurrentPage = maxPages;
    if (invCurrentPage < 1) invCurrentPage = 1;

    const startIndex = (invCurrentPage - 1) * invRowsPerPage;
    const paginatedItems = items.slice(startIndex, startIndex + invRowsPerPage);

    if (paginatedItems.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" class="empty-row"><strong>Nothing matches these filters</strong>Clear the search or press Reset filters.</td></tr>`;
        renderPagination(0, 0, 0);
        return;
    }

    const reservedMap = typeof getReservedSerials === 'function' ? getReservedSerials() : new Map();
    paginatedItems.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const localSNo = startIndex + idx + 1;
        const statusBadge = formatStatusBadge(item.status);
        const makeDisplay = item.make ? item.make.toUpperCase() : 'GENERIC';
        const makeClass = makeDisplay.toLowerCase().includes('gvr') ? 'make-gvr' : (makeDisplay.toLowerCase().includes('tokheim') ? 'make-tokheim' : (makeDisplay.toLowerCase().includes('florite') ? 'make-florite' : ''));
        const isDu = item.category.includes('DU');
        const reservedOn = reservedMap.get(safeLower(item.serial_no));

        const e = escapeHtml;
        tr.innerHTML = `
            <td class="text-muted">${localSNo}</td>
            <td>${isDu ? 'DU' : 'STP'}</td>
            <td><span class="make-pill ${makeClass}">${e(makeDisplay)}</span></td>
            <td>${e(item.model) || '—'}</td>
            <td>
                <span class="copyable-serial" data-copy="${e(item.serial_no)}" title="Copy serial">${e(item.serial_no) || '—'}<i class="fa-regular fa-copy"></i></span>
                ${reservedOn ? `<div style="margin-top:3px;"><span class="tag tag-reserved" title="Reserved on dispatch letter">Reserved · ${e(reservedOn)}</span></div>` : ''}
            </td>
            <td>${e(item.variant) || '—'}</td>
            <td style="white-space:nowrap;" data-label="Received">${e(formatDisplayDate(item.receipt_date)) || '—'}</td>
            <td>${statusBadge}</td>
            <td style="max-width:220px; font-size:0.86rem; color:var(--ink-2);" title="${e(item.remarks)}">${e(item.remarks) || ''}</td>
            <td>
                <div class="actions-cell">
                    <button class="btn-icon edit-btn" data-sno="${e(item.sno)}" data-cat="${isDu ? 'du' : 'stp'}" title="Change condition" aria-label="Change condition of ${e(item.serial_no)}">
                        <i class="fa-solid fa-pen"></i>
                    </button>
                    <button class="btn btn-outline btn-sm issue-gp-btn" data-sno="${e(item.sno)}" data-cat="${isDu ? 'DU' : 'STP'}" data-serial="${e(item.serial_no)}" title="Issue EMR for this asset">
                        Issue EMR
                    </button>
                </div>
            </td>
        `;

        tbody.appendChild(tr);
    });

    // Delegated event listener for Edit, Gate Pass, and Copy buttons in Inventory Table
    if (tableElem && !tableElem.dataset.listenerAttached) {
        tableElem.dataset.listenerAttached = 'true';
        tableElem.addEventListener('click', (e) => {
            const copyEl = e.target.closest('.copyable-serial');
            if (copyEl) {
                const textToCopy = copyEl.getAttribute('data-copy');
                copyToClipboard(textToCopy, 'Serial Number');
                return;
            }

            const editBtn = e.target.closest('.edit-btn');
            if (editBtn) {
                const sno = parseInt(editBtn.getAttribute('data-sno'), 10);
                const cat = editBtn.getAttribute('data-cat') || 'du';
                window.openStatusModal(cat, sno);
                return;
            }

            const gpBtn = e.target.closest('.issue-gp-btn');
            if (gpBtn) {
                const cat = gpBtn.getAttribute('data-cat') || 'DU';
                const serial = gpBtn.getAttribute('data-serial') || '';
                quickIssueGatePass(cat, serial);
                return;
            }
        });
    }

    renderPagination(totalItems, startIndex + 1, Math.min(startIndex + invRowsPerPage, totalItems));
}

function renderPagination(totalItems, startItem, endItem) {
    const container = document.getElementById('inventory-pagination');
    if (!container) return;

    if (totalItems === 0) {
        container.innerHTML = `
            <div class="pagination-left">
                <div class="pagination-info" id="inv-page-info">Showing 0 items</div>
            </div>
            <div class="pagination-controls"></div>
        `;
        return;
    }

    const maxPages = Math.max(Math.ceil(totalItems / invRowsPerPage), 1);
    if (invCurrentPage > maxPages) invCurrentPage = maxPages;
    if (invCurrentPage < 1) invCurrentPage = 1;

    let pagesHtml = '';
    const startPage = Math.max(1, invCurrentPage - 2);
    const endPage = Math.min(maxPages, invCurrentPage + 2);

    if (startPage > 1) {
        pagesHtml += `<button class="pagination-btn" data-page="1">1</button>`;
        if (startPage > 2) {
            pagesHtml += `<span style="padding: 0 4px; color: var(--text-secondary); line-height: 32px;">...</span>`;
        }
    }

    for (let p = startPage; p <= endPage; p++) {
        pagesHtml += `<button class="pagination-btn ${p === invCurrentPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
    }

    if (endPage < maxPages) {
        if (endPage < maxPages - 1) {
            pagesHtml += `<span style="padding: 0 4px; color: var(--text-secondary); line-height: 32px;">...</span>`;
        }
        pagesHtml += `<button class="pagination-btn" data-page="${maxPages}">${maxPages}</button>`;
    }

    container.innerHTML = `
        <div class="pagination-left">
            <div class="pagination-page-size">
                <span>Show:</span>
                <select id="inv-page-size">
                    <option value="10" ${invRowsPerPage === 10 ? 'selected' : ''}>10</option>
                    <option value="25" ${invRowsPerPage === 25 ? 'selected' : ''}>25</option>
                    <option value="50" ${invRowsPerPage === 50 ? 'selected' : ''}>50</option>
                    <option value="100" ${invRowsPerPage === 100 ? 'selected' : ''}>100</option>
                    <option value="all" ${invRowsPerPage > 1000 ? 'selected' : ''}>All</option>
                </select>
                <span>rows</span>
            </div>
            <div class="pagination-info" id="inv-page-info">Showing ${startItem}–${endItem} of ${totalItems} items</div>
        </div>
        <div class="pagination-controls">
            <button class="pagination-btn" id="inv-prev-page" ${invCurrentPage <= 1 ? 'disabled' : ''} title="Previous Page">
                <i class="fa-solid fa-chevron-left"></i>
            </button>
            ${pagesHtml}
            <button class="pagination-btn" id="inv-next-page" ${invCurrentPage >= maxPages ? 'disabled' : ''} title="Next Page">
                <i class="fa-solid fa-chevron-right"></i>
            </button>
        </div>
    `;

    const pageSizeSelect = document.getElementById('inv-page-size');
    if (pageSizeSelect) {
        pageSizeSelect.addEventListener('change', (e) => {
            invRowsPerPage = e.target.value === 'all' ? 999999 : (parseInt(e.target.value, 10) || 10);
            invCurrentPage = 1;
            renderInventoryList();
        });
    }

    const prevBtn = document.getElementById('inv-prev-page');
    if (prevBtn && invCurrentPage > 1) {
        prevBtn.addEventListener('click', () => {
            invCurrentPage--;
            renderInventoryList();
        });
    }

    const nextBtn = document.getElementById('inv-next-page');
    if (nextBtn && invCurrentPage < maxPages) {
        nextBtn.addEventListener('click', () => {
            invCurrentPage++;
            renderInventoryList();
        });
    }

    container.querySelectorAll('.pagination-btn[data-page]').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const pageNum = parseInt(e.currentTarget.getAttribute('data-page'), 10);
            if (!isNaN(pageNum) && pageNum !== invCurrentPage) {
                invCurrentPage = pageNum;
                renderInventoryList();
            }
        });
    });
}

// ==========================================================================
// Inward & Outward Ledgers
// ==========================================================================
function renderInwardLedger() {
    const table = document.getElementById('inward-table');
    const tbody = table ? table.querySelector('tbody') : null;
    if (!tbody) return;
    tbody.innerHTML = '';

    const globalSearch = safeLower(document.getElementById('global-search') ? document.getElementById('global-search').value : '');
    const inwardSearch = safeLower(document.getElementById('inward-search') ? document.getElementById('inward-search').value : '');
    const searchVal = inwardSearch || globalSearch;

    let items = (db.inword || []).map(normalizeRecord);

    if (searchVal !== '') {
        items = items.filter(i => {
            return safeLower(i.item).includes(searchVal) ||
                   safeLower(i.received_from).includes(searchVal) ||
                   safeLower(i.received_date).includes(searchVal);
        });
    }

    if (items.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="empty-row"><strong>No inward entries found</strong>Add an inward entry when material arrives.</td></tr>`;
        return;
    }

    items = [...items].sort((a, b) => safeDateTimestamp(b.received_date) - safeDateTimestamp(a.received_date));

    items.map(escapeRecord).forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="white-space:nowrap;">${formatDisplayDate(item.received_date) || '—'}</td>
            <td>${item.item || '—'}</td>
            <td>${item.received_from || '—'}</td>
            <td class="text-right serial">${item.qty || 1}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderOutwardLedger() {
    const globalSearch = safeLower(document.getElementById('global-search') ? document.getElementById('global-search').value : '');
    const outwardSearch = safeLower(document.getElementById('outward-search') ? document.getElementById('outward-search').value : '');
    const searchVal = outwardSearch || globalSearch;

    // 1. Render DU Outward (8 columns)
    const duTable = document.getElementById('outward-du-table');
    const duBody = duTable ? duTable.querySelector('tbody') : null;
    if (duBody) {
        duBody.innerHTML = '';
        let dus = (db.du_outword || []).map(normalizeRecord);
        if (searchVal) {
            dus = dus.filter(d => (safeLower(d.gate_pass_no).includes(searchVal) || safeLower(d.emr_no).includes(searchVal)) || safeLower(d.serial_no).includes(searchVal) || safeLower(d.destination).includes(searchVal) || safeLower(d.vendor).includes(searchVal) || safeLower(d.make).includes(searchVal));
        }
        dus.sort((a, b) => safeDateTimestamp(b.out_date || b.emr_date) - safeDateTimestamp(a.out_date || a.emr_date));
        if (dus.length === 0) {
            duBody.innerHTML = `<tr><td colspan="8" class="empty-row"><strong>No DU outward entries</strong>Dispatches recorded on EMRs appear here.</td></tr>`;
        } else {
            dus.map(escapeRecord).forEach(d => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="white-space:nowrap;">${formatDisplayDate(d.emr_date) || '—'}</td>
                    <td>${d.emr_no || '—'}</td>
                    <td><span class="serial">${d.serial_no || '—'}</span></td>
                    <td>${d.variant || d.type || 'DU'}</td>
                    <td><strong>${d.make || '—'}</strong> ${d.model || ''}</td>
                    <td>${d.destination || '—'}</td>
                    <td>${d.vendor || '—'}</td>
                    <td>${formatDisplayDate(d.out_date || d.emr_date) || '—'}</td>
                `;
                duBody.appendChild(tr);
            });
        }
    }

    // 2. Render STP Outward (8 columns)
    const stpTable = document.getElementById('outward-stp-table');
    const stpBody = stpTable ? stpTable.querySelector('tbody') : null;
    if (stpBody) {
        stpBody.innerHTML = '';
        let stps = (db.stp_outword || []).map(normalizeRecord);
        if (searchVal) {
            stps = stps.filter(s => (safeLower(s.gate_pass_no).includes(searchVal) || safeLower(s.emr_no).includes(searchVal)) || safeLower(s.serial_no).includes(searchVal) || safeLower(s.destination).includes(searchVal) || safeLower(s.vendor).includes(searchVal) || safeLower(s.make).includes(searchVal));
        }
        stps.sort((a, b) => safeDateTimestamp(b.out_date || b.emr_date) - safeDateTimestamp(a.out_date || a.emr_date));
        if (stps.length === 0) {
            stpBody.innerHTML = `<tr><td colspan="8" class="empty-row"><strong>No STP outward entries</strong>Dispatches recorded on EMRs appear here.</td></tr>`;
        } else {
            stps.map(escapeRecord).forEach(s => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="white-space:nowrap;">${formatDisplayDate(s.emr_date) || '—'}</td>
                    <td>${s.emr_no || '—'}</td>
                    <td><span class="serial">${s.serial_no || '—'}</span></td>
                    <td><span>${safeStr(s.hp_type || s.variant || 'STP').replace(/\s*HP$/i, '')} HP</span></td>
                    <td><strong>${s.make || '—'}</strong></td>
                    <td>${s.destination || '—'}</td>
                    <td>${s.vendor || '—'}</td>
                    <td>${formatDisplayDate(s.out_date || s.emr_date) || '—'}</td>
                `;
                stpBody.appendChild(tr);
            });
        }
    }

    // 3. Render Tank Outward (6 columns: #, EMR Date, EMR No, Capacity, Destination, Vendor)
    const tankTable = document.getElementById('outward-tank-table');
    const tankBody = tankTable ? tankTable.querySelector('tbody') : null;
    if (tankBody) {
        tankBody.innerHTML = '';
        let tanks = (db.tank_outword || []).map(normalizeRecord);
        if (searchVal) {
            tanks = tanks.filter(t => (safeLower(t.gate_pass_no).includes(searchVal) || safeLower(t.emr_no).includes(searchVal)) || safeLower(t.destination).includes(searchVal) || safeLower(t.vendor).includes(searchVal) || safeLower(t.tank_capacity).includes(searchVal));
        }
        tanks.sort((a, b) => safeDateTimestamp(b.out_date || b.emr_date) - safeDateTimestamp(a.out_date || a.emr_date));
        if (tanks.length === 0) {
            tankBody.innerHTML = `<tr><td colspan="6" class="empty-row"><strong>No tank outward entries</strong>Dispatches recorded on EMRs appear here.</td></tr>`;
        } else {
            tanks.map(escapeRecord).forEach((t, idx) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="text-muted">${t.sno || (idx + 1)}</td>
                    <td>${formatDisplayDate(t.emr_date) || '—'}</td>
                    <td>${t.emr_no || '—'}</td>
                    <td><strong>${t.tank_capacity || t.variant || '—'}</strong></td>
                    <td>${t.destination || '—'}</td>
                    <td>${t.vendor || '—'}</td>
                `;
                tankBody.appendChild(tr);
            });
        }
    }

    // 4. Render Others Outward (7 columns: #, EMR Date, EMR No, Material Description, Quantity, Destination, Vendor)
    const othTable = document.getElementById('outward-others-table');
    const othBody = othTable ? othTable.querySelector('tbody') : null;
    if (othBody) {
        othBody.innerHTML = '';
        let oths = (db.others_outword || []).map(normalizeRecord);
        if (searchVal) {
            oths = oths.filter(o => (safeLower(o.gate_pass_no).includes(searchVal) || safeLower(o.emr_no).includes(searchVal)) || safeLower(o.material).includes(searchVal) || safeLower(o.destination).includes(searchVal) || safeLower(o.vendor).includes(searchVal));
        }
        oths.sort((a, b) => safeDateTimestamp(b.out_date || b.emr_date) - safeDateTimestamp(a.out_date || a.emr_date));
        if (oths.length === 0) {
            othBody.innerHTML = `<tr><td colspan="7" class="empty-row"><strong>No other outward entries</strong>Dispatches recorded on EMRs appear here.</td></tr>`;
        } else {
            oths.map(escapeRecord).forEach((o, idx) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="text-muted">${o.sno || (idx + 1)}</td>
                    <td>${formatDisplayDate(o.emr_date) || '—'}</td>
                    <td>${o.emr_no || '—'}</td>
                    <td>${o.material || o.item || '—'}</td>
                    <td class="serial">${o.qty || 1}</td>
                    <td>${o.destination || '—'}</td>
                    <td>${o.vendor || '—'}</td>
                `;
                othBody.appendChild(tr);
            });
        }
    }
}

// ==========================================================================
// Modals & Form Submission Management
// ==========================================================================
function setupModals() {
    const recordModal = document.getElementById('record-modal');
    const statusModal = document.getElementById('status-modal');
    const addEntryBtn = document.getElementById('add-entry-btn');
    const addInwardBtn = document.getElementById('add-inward-btn');
    const addOutwardBtn = document.getElementById('add-outward-btn');
    const outCategory = document.getElementById('out-category');

    if (outCategory) {
        outCategory.addEventListener('change', (e) => {
            const cat = e.target.value;
            const serialGroup = document.getElementById('out-serial-group');
            const detailsRow = document.getElementById('out-details-row');
            const qtyGroup = document.getElementById('out-qty-group');
            const serialLabel = serialGroup ? serialGroup.querySelector('label') : null;
            const serialInput = document.getElementById('out-serial');
            const variantLabel = detailsRow ? detailsRow.querySelector('label[for="out-variant"]') : null;
            const variantInput = document.getElementById('out-variant');

            if (cat === 'du' || cat === 'stp') {
                if (serialGroup) serialGroup.style.display = 'block';
                if (serialLabel) serialLabel.textContent = 'Serial No *';
                if (serialInput) serialInput.placeholder = 'Asset serial no';
                if (detailsRow) detailsRow.style.display = 'flex';
                if (variantLabel) variantLabel.textContent = (cat === 'du') ? 'Variant / Specs' : 'HP Rating (0.75 / 1.5 HP)';
                if (variantInput) variantInput.placeholder = (cat === 'du') ? 'e.g. MPD - 2P4N4D' : 'e.g. 0.75 or 1.5';
                if (qtyGroup) {
                    qtyGroup.style.display = 'none';
                    document.getElementById('out-qty').value = 1;
                }
            } else if (cat === 'tank') {
                if (serialGroup) serialGroup.style.display = 'none';
                if (detailsRow) detailsRow.style.display = 'flex';
                if (variantLabel) variantLabel.textContent = 'Tank Capacity *';
                if (variantInput) variantInput.placeholder = 'e.g. 16KL or 20KL';
                if (qtyGroup) {
                    qtyGroup.style.display = 'none';
                    document.getElementById('out-qty').value = 1;
                }
            } else if (cat === 'others') {
                if (serialGroup) serialGroup.style.display = 'block';
                if (serialLabel) serialLabel.textContent = 'Material / Item Description *';
                if (serialInput) serialInput.placeholder = 'e.g. Submersible Cable, Flange, Pipe';
                if (detailsRow) detailsRow.style.display = 'none';
                if (qtyGroup) qtyGroup.style.display = 'block';
            }
        });
    }

    // Open Add Record Modal
    if (addEntryBtn) {
        addEntryBtn.addEventListener('click', () => {
            const form = document.getElementById('record-form');
            if (form) form.reset();
            
            if (outCategory) {
                outCategory.value = 'du';
                outCategory.dispatchEvent(new Event('change'));
            }
            
            const today = todayIso();
            const invRec = document.getElementById('inv-receipt-date');
            const inDate = document.getElementById('in-date');
            const outEmr = document.getElementById('out-emr-date');
            const outDate = document.getElementById('out-date');
            if (invRec) invRec.value = today;
            if (inDate) inDate.value = today;
            if (outEmr) outEmr.value = today;
            if (outDate) outDate.value = today;

            if (recordModal) recordModal.classList.add('active');
            const defaultTab = document.querySelector('.modal-tab-btn[data-type="inventory"]');
            if (defaultTab) defaultTab.click();
        });
    }

    if (addInwardBtn) {
        addInwardBtn.addEventListener('click', () => {
            if (addEntryBtn) addEntryBtn.click();
            const inwardTab = document.querySelector('.modal-tab-btn[data-type="inward"]');
            if (inwardTab) inwardTab.click();
        });
    }

    if (addOutwardBtn) {
        addOutwardBtn.addEventListener('click', () => {
            if (addEntryBtn) addEntryBtn.click();
            const outwardTab = document.querySelector('.modal-tab-btn[data-type="outward"]');
            if (outwardTab) outwardTab.click();
        });
    }

    // Close buttons
    const recClose = document.getElementById('record-modal-close');
    const recCancel = document.getElementById('record-modal-cancel');
    const statClose = document.getElementById('status-modal-close');
    const statCancel = document.getElementById('status-modal-cancel');

    if (recClose && recordModal) recClose.addEventListener('click', () => recordModal.classList.remove('active'));
    if (recCancel && recordModal) recCancel.addEventListener('click', () => recordModal.classList.remove('active'));
    if (statClose && statusModal) statClose.addEventListener('click', () => statusModal.classList.remove('active'));
    if (statCancel && statusModal) statCancel.addEventListener('click', () => statusModal.classList.remove('active'));

    // Unified background overlay click handler for all modals
    window.addEventListener('click', (e) => {
        if (e.target && e.target.classList && e.target.classList.contains('modal-overlay')) {
            e.target.classList.remove('active');
        }
    });

    // Any element with data-close-modal closes the modal it sits in
    document.addEventListener('click', (e) => {
        const closer = e.target.closest('[data-close-modal]');
        if (closer) {
            const overlay = closer.closest('.modal-overlay');
            if (overlay) overlay.classList.remove('active');
        }
    });

    // Keyboard ESC key closes any active modal
    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal-overlay.active').forEach(m => m.classList.remove('active'));
        }
    });

    // Modal internal tab selector
    const tabBtns = document.querySelectorAll('.modal-tab-btn');
    const fieldSections = document.querySelectorAll('.modal-field-section');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.getAttribute('data-type');
            
            tabBtns.forEach(t => t.classList.remove('active'));
            btn.classList.add('active');

            fieldSections.forEach(section => {
                if (section.id === `fields-${type}`) {
                    section.classList.add('active');
                } else {
                    section.classList.remove('active');
                }
            });

            // Adjust input validations dynamically based on active tab
            const invMake = document.getElementById('inv-make');
            const invSerial = document.getElementById('inv-serial');
            const inItem = document.getElementById('in-item');
            const inFrom = document.getElementById('in-from');
            const inQty = document.getElementById('in-qty');
            const outEmrNo = document.getElementById('out-emr-no');
            const outEmrDate = document.getElementById('out-emr-date');
            const outDestination = document.getElementById('out-destination');

            if (invMake) invMake.required = (type === 'inventory');
            if (invSerial) invSerial.required = (type === 'inventory');
            if (inItem) inItem.required = (type === 'inward');
            if (inFrom) inFrom.required = (type === 'inward');
            if (inQty) inQty.required = (type === 'inward');
            if (outEmrNo) outEmrNo.required = (type === 'outward');
            if (outEmrDate) outEmrDate.required = (type === 'outward');
            if (outDestination) outDestination.required = (type === 'outward');
        });
    });

    // Form Submissions
    const recordForm = document.getElementById('record-form');
    const statusForm = document.getElementById('status-form');
    if (recordForm) recordForm.addEventListener('submit', handleNewRecord);
    if (statusForm) statusForm.addEventListener('submit', handleStatusUpdate);
}

// Active Status Edit Modal Handler
function openStatusModal(category, sno) {
    const form = document.getElementById('status-form');
    if (form) form.reset();
    
    const catInput = document.getElementById('edit-item-category');
    const snoInput = document.getElementById('edit-item-sno');
    if (catInput) catInput.value = category;
    if (snoInput) snoInput.value = sno;

    let item = null;
    const badge = document.getElementById('edit-item-category-badge');
    const isDu = String(category).toLowerCase() === 'du';
    
    if (isDu) {
        item = (db.inventory.pressure_dus || []).find(d => Number(d.sno) === Number(sno));
        if (badge) {
            badge.className = 'badge badge-new';
            badge.textContent = 'Dispenser';
        }
    } else {
        item = (db.inventory.stps || []).find(s => Number(s.sno) === Number(sno));
        if (badge) {
            badge.className = 'badge badge-info';
            badge.textContent = 'STP';
        }
    }

    if (item) {
        const titleElem = document.getElementById('edit-item-title');
        const statusElem = document.getElementById('edit-item-status');
        const remarksElem = document.getElementById('edit-item-remarks');
        if (titleElem) titleElem.textContent = `${item.make} ${item.model || ''} - S/N: ${item.serial_no}`;
        if (statusElem) {
            const st = safeUpper(item.status);
            statusElem.value = (st === 'U' || st.includes('USED')) ? 'U'
                : (st === 'S' || st.includes('UNSERVICE') || st.includes('SCRAP')) ? 'S' : 'N';
        }
        if (remarksElem) remarksElem.value = item.remarks || '';
        const modal = document.getElementById('status-modal');
        if (modal) modal.classList.add('active');
    }
}

function handleStatusUpdate(e) {
    e.preventDefault();
    const catVal = safeLower(document.getElementById('edit-item-category') ? document.getElementById('edit-item-category').value : '');
    const snoVal = parseInt(document.getElementById('edit-item-sno') ? document.getElementById('edit-item-sno').value : 0, 10);
    const newStatus = document.getElementById('edit-item-status') ? document.getElementById('edit-item-status').value : 'N';
    const newRemarks = document.getElementById('edit-item-remarks') ? document.getElementById('edit-item-remarks').value : '';

    let item = null;
    if (catVal === 'du') {
        item = (db.inventory.pressure_dus || []).find(d => Number(d.sno) === snoVal);
    } else {
        item = (db.inventory.stps || []).find(s => Number(s.sno) === snoVal);
    }

    if (item) {
        item.status = newStatus;
        item.remarks = newRemarks;
        saveState();
        refreshAllViews();
        const modal = document.getElementById('status-modal');
        if (modal) modal.classList.remove('active');
        showToast("Item condition status updated successfully!");
    }
}

function handleNewRecord(e) {
    e.preventDefault();
    
    const activeTabBtn = document.querySelector('.modal-tab-btn.active');
    const transactionType = activeTabBtn ? activeTabBtn.getAttribute('data-type') : 'inventory';
    
    if (transactionType === 'inventory') {
        const category = document.getElementById('inv-category').value;
        const make = safeUpper(document.getElementById('inv-make').value);
        const model = safeStr(document.getElementById('inv-model').value);
        const serial = safeStr(document.getElementById('inv-serial').value);
        const variant = safeStr(document.getElementById('inv-variant').value);
        const status = document.getElementById('inv-status').value;
        const receiptDate = formatDisplayDate(document.getElementById('inv-receipt-date').value);
        const remarks = safeStr(document.getElementById('inv-remarks').value);

        const allInStore = [...(db.inventory.pressure_dus || []), ...(db.inventory.stps || [])];
        if (allInStore.some(x => safeLower(x.serial_no) === safeLower(serial))) {
            alert(`Serial Number '${serial}' is already in In-Store Inventory. Use the Edit button on that row to change its condition instead.`);
            return;
        }

        if (category === 'du') {
            const nextSno = (db.inventory.pressure_dus || []).reduce((max, d) => Number(d.sno) > max ? Number(d.sno) : max, 0) + 1;
            db.inventory.pressure_dus.push({
                sno: nextSno,
                type: 'Pressure',
                variant: variant || 'Dispenser',
                make, model, serial_no: serial, receipt_date: receiptDate, status, remarks
            });
        } else {
            const nextSno = (db.inventory.stps || []).reduce((max, s) => Number(s.sno) > max ? Number(s.sno) : max, 0) + 1;
            db.inventory.stps.push({
                sno: nextSno,
                hp_type: variant || '0.75',
                make, model, serial_no: serial, receipt_date: receiptDate, status, remarks
            });
        }
        showToast("New asset added to inventory!");
        
    } else if (transactionType === 'inward') {
        const item = safeStr(document.getElementById('in-item').value);
        const receivedFrom = safeStr(document.getElementById('in-from').value);
        const qty = parseInt(document.getElementById('in-qty').value, 10) || 1;
        const receivedDate = formatDisplayDate(document.getElementById('in-date').value);

        db.inword.push({
            item, received_from: receivedFrom, received_date: receivedDate, qty
        });

        showToast("Inbound receipt ledger updated!");

    } else if (transactionType === 'outward') {
        const cat = document.getElementById('out-category').value;
        const gpNo = safeStr(document.getElementById('out-emr-no').value);
        const gpDate = formatDisplayDate(document.getElementById('out-emr-date').value);
        const destination = safeStr(document.getElementById('out-destination').value);
        const vendor = safeStr(document.getElementById('out-vendor').value);
        const qty = parseInt(document.getElementById('out-qty').value, 10) || 1;
        const outDate = formatDisplayDate(document.getElementById('out-date').value) || gpDate;

        if (cat === 'du' || cat === 'stp') {
            const ser = safeStr(document.getElementById('out-serial').value);
            if (!ser) {
                alert('Please enter the Serial No of the asset being dispatched.');
                return;
            }
            const list = cat === 'du' ? db.inventory.pressure_dus : db.inventory.stps;
            if (!list.some(x => safeLower(x.serial_no) === safeLower(ser)) &&
                !confirm(`Serial '${ser}' is not in In-Store ${cat === 'du' ? 'DU' : 'STP'} inventory.\n\nRecord the outward entry anyway?`)) {
                return;
            }
        }

        if (cat === 'du') {
            const serial = safeStr(document.getElementById('out-serial').value);
            const variant = safeStr(document.getElementById('out-variant').value);
            const make = safeUpper(document.getElementById('out-make').value);
            const model = document.getElementById('out-model') ? safeStr(document.getElementById('out-model').value) : '';
            const stock = db.inventory.pressure_dus.find(d => safeLower(d.serial_no) === safeLower(serial)) || {};

            const nextSno = (db.du_outword || []).reduce((max, d) => Number(d.sno) > max ? Number(d.sno) : max, 0) + 1;
            db.du_outword.push({
                sno: nextSno,
                type: stock.type || 'Pressure',
                variant: variant || stock.variant || 'MPD',
                make: make || safeUpper(stock.make), model: model || stock.model || '', serial_no: stock.serial_no || serial,
                status: stock.status || '', receipt_date: stock.receipt_date || '',
                gate_pass_date: gpDate, gate_pass_no: gpNo,
                emr_date: gpDate, emr_no: gpNo,
                destination, vendor, out_date: outDate
            });

            const idx = db.inventory.pressure_dus.findIndex(d => safeLower(d.serial_no) === safeLower(serial));
            if (idx !== -1) {
                db.inventory.pressure_dus.splice(idx, 1);
            }

        } else if (cat === 'stp') {
            const serial = safeStr(document.getElementById('out-serial').value);
            const variant = safeStr(document.getElementById('out-variant').value);
            const make = safeUpper(document.getElementById('out-make').value);
            const stock = db.inventory.stps.find(s => safeLower(s.serial_no) === safeLower(serial)) || {};

            const nextSno = (db.stp_outword || []).reduce((max, s) => Number(s.sno) > max ? Number(s.sno) : max, 0) + 1;
            db.stp_outword.push({
                sno: nextSno,
                hp_type: safeStr(variant || stock.hp_type || '0.75').replace(/\s*HP$/i, ''),
                make: make || safeUpper(stock.make), model: stock.model || '', serial_no: stock.serial_no || serial,
                status: stock.status || '', receipt_date: stock.receipt_date || '',
                gate_pass_date: gpDate, gate_pass_no: gpNo,
                emr_date: gpDate, emr_no: gpNo,
                destination, vendor, out_date: outDate
            });

            const idx = db.inventory.stps.findIndex(s => safeLower(s.serial_no) === safeLower(serial));
            if (idx !== -1) {
                db.inventory.stps.splice(idx, 1);
            }

        } else if (cat === 'tank') {
            const capacity = safeStr(document.getElementById('out-variant').value);
            const nextSno = (db.tank_outword || []).reduce((max, t) => Number(t.sno) > max ? Number(t.sno) : max, 0) + 1;
            db.tank_outword.push({
                sno: nextSno,
                tank_capacity: capacity,
                gate_pass_date: gpDate, gate_pass_no: gpNo,
                emr_date: gpDate, emr_no: gpNo,
                destination, vendor
            });

        } else if (cat === 'others') {
            const desc = safeStr(document.getElementById('out-serial').value) || safeStr(document.getElementById('out-variant').value);
            const nextSno = (db.others_outword || []).reduce((max, o) => Number(o.sno) > max ? Number(o.sno) : max, 0) + 1;
            db.others_outword.push({
                sno: nextSno,
                material: desc, qty,
                gate_pass_date: gpDate, gate_pass_no: gpNo,
                emr_date: gpDate, emr_no: gpNo,
                destination, vendor
            });
        }
        showToast("Outward Gate Pass transaction registered successfully!");
    }

    saveState();
    refreshAllViews();
    populateFilterDropdowns();
    const modal = document.getElementById('record-modal');
    if (modal) modal.classList.remove('active');
}

// ==========================================================================
// Data Actions (Export, Baseline Reset, Theme Toggle)
// ==========================================================================
function setupDataActions() {
    // 1. Export JSON
    const btnExportJson = document.getElementById('btn-export-json');
    if (btnExportJson) {
        btnExportJson.addEventListener('click', async () => {
            const photoCount = (db.gate_passes || []).reduce((n, g) => n + (Array.isArray(g.photos) ? g.photos.length : 0), 0);
            const fileName = `hpcl_stores_backup_${todayLocal().replace(/\//g, '-')}.json`;
            if (photoCount > 0 && typeof exportPhotosForBackup === 'function' &&
                confirm(`Include the ${photoCount} dispatch photo(s) stored on this device?\n\nOK = include photos (bigger file)\nCancel = data only`)) {
                const photoData = await exportPhotosForBackup();
                downloadTextFile(JSON.stringify({ ...db, _photo_data: photoData }), fileName, 'application/json');
                showToast(`Backup downloaded with ${Object.keys(photoData).length} photo(s).`);
                return;
            }
            downloadTextFile(JSON.stringify(db, null, 2), fileName, 'application/json');
            showToast("JSON database backup downloaded!");
        });
    }

    // 1b. Restore JSON backup
    const btnImportJson = document.getElementById('btn-import-json');
    const importJsonFile = document.getElementById('import-json-file');
    if (btnImportJson && importJsonFile) {
        btnImportJson.addEventListener('click', () => importJsonFile.click());
        importJsonFile.addEventListener('change', (e) => {
            const file = e.target.files && e.target.files[0];
            e.target.value = '';
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async () => {
                let data;
                try {
                    data = JSON.parse(reader.result);
                } catch (err) {
                    alert('This file is not valid JSON: ' + err.message);
                    return;
                }
                if (!data || typeof data !== 'object' || !data.inventory) {
                    alert('This does not look like an HPCL Inventory backup (no "inventory" section found).');
                    return;
                }
                const nDu = (data.inventory.pressure_dus || []).length;
                const nStp = (data.inventory.stps || []).length;
                const nEmr = (data.gate_passes || []).length;
                if (!confirm(`Restore backup '${file.name}'?\n\n${nDu} DUs, ${nStp} STPs, ${nEmr} EMRs.\n\nThis replaces all data currently on this device.`)) return;
                const photoData = data._photo_data;
                delete data._photo_data;
                db = data;
                standardizeDatabase();
                saveState();
                refreshAllViews();
                if (photoData && typeof importPhotosFromBackup === 'function') {
                    importPhotosFromBackup(photoData).then(n => {
                        refreshAllViews();
                        showToast(`Backup restored with ${n} photo(s).`);
                    });
                } else {
                    showToast('Backup restored.');
                }
            };
            reader.onerror = () => alert('Could not read the file.');
            reader.readAsText(file);
        });
    }

    // 2. Export Inventory CSV (Full)
    const btnExportCsv = document.getElementById('btn-export-csv');
    if (btnExportCsv) {
        btnExportCsv.addEventListener('click', () => exportCollectionToCsv('inventory_all'));
    }

    // 3. Export Filtered Inventory View
    const btnExportFiltered = document.getElementById('btn-export-filtered-inv');
    if (btnExportFiltered) {
        btnExportFiltered.addEventListener('click', () => exportCollectionToCsv('inventory_filtered'));
    }

    // 4. Export Inward
    const btnExportInward = document.getElementById('btn-export-inward');
    if (btnExportInward) {
        btnExportInward.addEventListener('click', () => exportCollectionToCsv('inward'));
    }

    // 5. Export Outward
    const btnExportOutward = document.getElementById('btn-export-outward');
    if (btnExportOutward) {
        btnExportOutward.addEventListener('click', () => exportCollectionToCsv('outward'));
    }

    // 6. Export Gate Passes
    const btnExportGps = document.getElementById('btn-export-gps');
    if (btnExportGps) {
        btnExportGps.addEventListener('click', () => exportCollectionToCsv('gate_passes'));
    }

    // 7. Reset Data
    const btnResetData = document.getElementById('btn-reset-data');
    if (btnResetData) {
        btnResetData.addEventListener('click', () => {
            if (confirm("Are you sure you want to restore the inventory baseline? This will delete all your local entries and modifications!")) {
                localStorage.removeItem('hpcl_inventory_db');
                loadDatabase();
                showToast("Database restored to baseline!");
            }
        });
    }

    // 8. Theme Toggle
    const themeBtn = document.getElementById('theme-toggle-btn');
    if (themeBtn) {
        themeBtn.addEventListener('click', () => {
            const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
            const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
            
            document.documentElement.setAttribute('data-theme', newTheme);
            localStorage.setItem('hpcl_theme', newTheme);
            
            const btnSpan = themeBtn.querySelector('span');
            if (btnSpan) {
                btnSpan.textContent = newTheme === 'dark' ? 'Light mode' : 'Dark mode';
            }

            updateCharts();
        });
    }
}

// Trigger a file download from a string (Blob URL: no size limit, and '#' or '%' in data can't truncate it)
function downloadTextFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function csvCell(val) {
    let v = safeStr(val);
    // Keep Excel from evaluating cells that start with = + - @ as formulas
    if (/^[=+\-@]/.test(v) && !/^-?\d+(\.\d+)?$/.test(v)) v = "'" + v;
    return `"${v.replace(/"/g, '""')}"`;
}

function csvRow(cells) {
    return cells.map(csvCell).join(',') + '\r\n';
}

function stpHp(rec) {
    const hp = safeStr(rec.hp_type || rec.variant).replace(/\s*HP$/i, '');
    return hp ? `${hp} HP` : '';
}

function exportCollectionToCsv(type) {
    const today = todayLocal().replace(/\//g, '-');
    let fileName = `hpcl_export_${today}.csv`;
    let csv = '';

    if (type === 'inventory_all' || type === 'inventory_filtered') {
        fileName = type === 'inventory_filtered' ? `hpcl_inventory_view_${today}.csv` : `hpcl_inventory_${today}.csv`;
        csv += csvRow(['S.No', 'Category', 'Make', 'Model', 'Serial No', 'Variant / HP', 'Receipt Date', 'Condition', 'Remarks']);

        let items;
        if (type === 'inventory_filtered') {
            // Exactly what the In-Store table shows (same filters, search and sort)
            items = getInventoryViewItems().map(i => ({
                cat: i.category && i.category.includes('STP') ? 'STP' : 'Dispenser',
                make: i.make, model: i.model, serial: i.serial_no, variant: i.variant,
                date: i.receipt_date, status: i.status, remarks: i.remarks
            }));
        } else {
            items = [];
            (db.inventory.pressure_dus || []).forEach(d => {
                items.push({ cat: 'Dispenser', make: d.make, model: d.model, serial: d.serial_no, variant: d.variant, date: d.receipt_date, status: d.status, remarks: d.remarks });
            });
            (db.inventory.stps || []).forEach(st => {
                items.push({ cat: 'STP', make: st.make, model: st.model, serial: st.serial_no, variant: stpHp(st), date: st.receipt_date, status: st.status, remarks: st.remarks });
            });
        }
        items.forEach((d, idx) => {
            csv += csvRow([idx + 1, d.cat, d.make, d.model, d.serial, d.variant, formatDisplayDate(d.date), d.status, d.remarks]);
        });

    } else if (type === 'inward') {
        fileName = `hpcl_inward_ledger_${today}.csv`;
        csv += csvRow(['Received Date', 'Material / Item', 'Received From', 'Quantity']);
        (db.inword || []).forEach(i => {
            csv += csvRow([formatDisplayDate(i.received_date), i.item, i.received_from, i.qty || 1]);
        });

    } else if (type === 'outward') {
        fileName = `hpcl_outward_ledger_${today}.csv`;
        csv += csvRow(['Category', 'EMR Date', 'EMR No', 'Gate Pass No', 'Serial / Capacity', 'Make / Specs', 'Destination', 'Vendor', 'Out Date']);
        (db.du_outword || []).forEach(d => {
            csv += csvRow(['DU', formatDisplayDate(d.emr_date || d.gate_pass_date), d.emr_no || d.gate_pass_no, d.gate_pass_no, d.serial_no, `${safeStr(d.make)} ${safeStr(d.model)}`.trim(), d.destination, d.vendor, formatDisplayDate(d.out_date)]);
        });
        (db.stp_outword || []).forEach(st => {
            csv += csvRow(['STP', formatDisplayDate(st.emr_date || st.gate_pass_date), st.emr_no || st.gate_pass_no, st.gate_pass_no, st.serial_no, `${safeStr(st.make)} ${stpHp(st) ? '(' + stpHp(st) + ')' : ''}`.trim(), st.destination, st.vendor, formatDisplayDate(st.out_date)]);
        });
        (db.tank_outword || []).forEach(t => {
            csv += csvRow(['Tank', formatDisplayDate(t.emr_date || t.gate_pass_date), t.emr_no || t.gate_pass_no, t.gate_pass_no, t.tank_capacity, 'Storage Tank', t.destination, t.vendor, formatDisplayDate(t.out_date)]);
        });
        (db.others_outword || []).forEach(o => {
            csv += csvRow(['Others', formatDisplayDate(o.emr_date || o.gate_pass_date), o.emr_no || o.gate_pass_no, o.gate_pass_no, `Qty: ${o.qty || 1}`, o.material, o.destination, o.vendor, formatDisplayDate(o.out_date)]);
        });

    } else if (type === 'letters') {
        fileName = `hpcl_dispatch_letters_${today}.csv`;
        csv += csvRow(['Letter No', 'Letter Date', 'Status', 'Destination', 'Vendor', 'Purpose', 'Serial No(s)', 'Other Material', 'Issued By', 'Designation', 'Office Ref', 'EMR No', 'Dispatched On', 'Vehicle No', 'Dispatched By']);
        (db.dispatch_letters || []).forEach(l => {
            csv += csvRow([l.dl_no, formatDisplayDate(l.dl_date), l.status, l.destination, l.vendor, l.purpose,
                (l.items || []).map(i => `${i.category} ${i.serial_no}`).join(', '),
                (l.other_items || []).map(o => `${o.material} x${o.qty}`).join(', '),
                l.issued_by, l.designation, l.ref, l.emr_no, formatDisplayDate(l.dispatched_on), l.vehicle_no, l.dispatched_by]);
        });

    } else if (type === 'gate_passes' || type === 'emr' || type === 'emrs') {
        fileName = `hpcl_emr_vouchers_${today}.csv`;
        csv += csvRow(['EMR No', 'Out Date', 'Dispatch Letter', 'Category', 'No. of Assets', 'Serial No(s)', 'Make', 'Model', 'Vendor', 'Destination', 'Gate Pass No', 'Vehicle No', 'Driver', 'Dispatched By', 'Photos', 'Remarks']);
        (db.gate_passes || []).forEach(g => {
            const items = Array.isArray(g.items) && g.items.length ? g.items : null;
            const serials = items ? items.map(i => i.serial_no).join(', ') : g.serial_no;
            const photos = Array.isArray(g.photos) ? g.photos : [];
            const photoLinks = photos.map(p => p.drive_url).filter(Boolean).join(' ');
            csv += csvRow([g.emr_no || g.gate_pass_no, formatDisplayDate(g.out_date), g.dl_no, g.category, items ? items.length : 1, serials, g.make, g.model, g.vendor, g.destination, g.gate_pass_no, g.vehicle_no, g.driver, g.dispatched_by, photoLinks || (photos.length ? `${photos.length} (on device)` : ''), g.remarks]);
        });
    }

    // UTF-8 BOM so Excel shows ₹ / Hindi text correctly
    downloadTextFile('﻿' + csv, fileName, 'text/csv;charset=utf-8');
    showToast(`CSV exported: ${fileName}`);
}

function showToast(message) {
    document.querySelectorAll('.toast-notification').forEach(t => t.remove());
    const toast = document.createElement('div');
    toast.className = 'toast-notification';
    toast.setAttribute('role', 'status');
    toast.textContent = safeStr(message);
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.style.animation = 'toastSlideOut 0.25s forwards';
        setTimeout(() => toast.remove(), 250);
    }, 3200);
}

const styleSheet = document.createElement("style");
styleSheet.innerText = `
@keyframes toastSlideIn {
    from { opacity: 0; transform: translateY(30px) scale(0.9); }
    to { opacity: 1; transform: translateY(0) scale(1); }
}
@keyframes toastSlideOut {
    from { opacity: 1; transform: translateY(0); }
    to { opacity: 0; transform: translateY(15px); }
}
`;
document.head.appendChild(styleSheet);

// ==========================================================================
// Excel Import Logic via SheetJS
// ==========================================================================
function setupExcelImport() {
    const importModal = document.getElementById('import-modal');
    const importBtn = document.getElementById('import-data-btn');
    const importClose = document.getElementById('import-modal-close');
    const fileDragArea = document.getElementById('file-drag-area');
    const fileInput = document.getElementById('import-excel-file');
    const loadingIndicator = document.getElementById('import-loading-indicator');

    if (!importBtn) return;

    // DU Template Download
    const downloadTemplateBtn = document.getElementById('download-template-btn');
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener('click', (e) => {
            e.preventDefault();
            downloadDuExcelTemplate();
        });
    }

    // STP Template Download
    const downloadStpTemplateBtn = document.getElementById('download-stp-template-btn');
    if (downloadStpTemplateBtn) {
        downloadStpTemplateBtn.addEventListener('click', (e) => {
            e.preventDefault();
            downloadStpExcelTemplate();
        });
    }

    function downloadDuExcelTemplate() {
        if (typeof XLSX === 'undefined') {
            alert("SheetJS library is loading. Please try again in a few seconds.");
            return;
        }
        const wb = XLSX.utils.book_new();
        const rows = [
            [
                "SNO.",
                "TYPE OF DUs",
                "VARIANT (SDM/SDD)/(HDM/HDD)/(MPD)/(SDHD)",
                "MAKE (GVR/MIDCO/TOKHEIM/any other)",
                "MODEL",
                "SERIAL NO. OF THE DISPENSING UNITS",
                "(DD/MM/YYYY) -DATE OF RECEIPT AT STORAGE POINT",
                "-NEW(N)/USED(U)/ UNSERVICEBLE(S)",
                "REMARKS"
            ],
            [1, "Pressure", "MPD - 2P4N2D", "TOKHEIM", "TIQ2442SHPN-Q330", "M2350433", "22/01/2024", "U", "Sample used dispenser remarks"],
            [2, "Pressure", "MPD - 2P4N4D", "GVR", "LH2PMU2TNEA3E32", "202412000642", "19/12/2024", "N", "Sample new dispenser remarks"],
            [3, "Pressure", "MPD - 2P4N4D", "MIDCO", "MIDCO-2024-MPD", "MD24001928", "15/11/2024", "N", "Sample new Midco dispenser remarks"]
        ];

        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!cols'] = [
            { wch: 8 },
            { wch: 14 },
            { wch: 22 },
            { wch: 14 },
            { wch: 22 },
            { wch: 24 },
            { wch: 18 },
            { wch: 12 },
            { wch: 30 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, "Inventory");
        XLSX.writeFile(wb, "hpcl_du_inventory_template.xlsx");
        showToast("DU Inventory Excel Template downloaded!");
    }

    function downloadStpExcelTemplate() {
        if (typeof XLSX === 'undefined') {
            alert("SheetJS library is loading. Please try again in a few seconds.");
            return;
        }
        const wb = XLSX.utils.book_new();
        const rows = [
            [
                "SNO.",
                "TYPE OF STPs(0.75 HP/1.5 HP)",
                "MAKE (GVR/KBL/Florite)",
                "MAKE (GVR/Flo-rite/KBL/any other)",
                "MODEL",
                "(DD/MM/YYYY) - DATE OF RECEIPT",
                "NEW(N)/USED(U)/ UNSERVICEBLE(S)",
                "REMARKS"
            ],
            [1, 0.75, "GVR", "RJ013261I", "NA", "16/07/2025", "N", "Sample STP remarks"],
            [2, 1.5, "GVR", "072163I", "NA", "19/07/2025", "N", "Another sample STP remarks"],
            [3, 0.75, "FLORITE", "FL2025-0091", "NA", "01/08/2025", "U", "Sample used Florite pump"]
        ];

        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!cols'] = [
            { wch: 8 },
            { wch: 28 },
            { wch: 22 },
            { wch: 26 },
            { wch: 12 },
            { wch: 20 },
            { wch: 22 },
            { wch: 30 }
        ];
        XLSX.utils.book_append_sheet(wb, ws, "Inventory");
        XLSX.writeFile(wb, "hpcl_stp_inventory_template.xlsx");
        showToast("STP Inventory Excel Template downloaded!");
    }

    importBtn.addEventListener('click', () => {
        if (importModal) importModal.classList.add('active');
        if (loadingIndicator) loadingIndicator.style.display = 'none';
        if (fileDragArea) fileDragArea.style.display = 'block';
    });

    if (importClose && importModal) {
        importClose.addEventListener('click', () => {
            importModal.classList.remove('active');
        });
    }

    if (fileDragArea && fileInput) {
        fileDragArea.addEventListener('click', () => {
            fileInput.click();
        });

        fileDragArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            fileDragArea.classList.add('drag-over');
        });

        fileDragArea.addEventListener('dragleave', () => {
            fileDragArea.classList.remove('drag-over');
        });

        fileDragArea.addEventListener('drop', (e) => {
            e.preventDefault();
            fileDragArea.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleExcelFile(files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            const files = e.target.files;
            if (files.length > 0) {
                handleExcelFile(files[0]);
            }
            // Reset so choosing the same file again still triggers an import
            e.target.value = '';
        });
    }

    function parseDateValue(rawDate) {
        if (rawDate === null || rawDate === undefined || rawDate === '') return null;
        if (typeof rawDate === 'string') {
            const low = rawDate.trim().toLowerCase();
            if (!low || ['nil', 'na', '-', 'none', 'null'].includes(low)) return null;
        }
        const norm = normalizeDateValue(rawDate);
        return norm || null;
    }

    function normalizeStatusValue(val) {
        if (!val) return 'N';
        const s = String(val).trim().toUpperCase();
        if (s.startsWith('U') || s.includes('USED')) return 'U';
        if (s.startsWith('S') || s.includes('SCRAP') || s.includes('UNSERVICE')) return 'S';
        if (s.startsWith('N') || s.includes('NEW')) return 'N';
        return s.charAt(0) || 'N';
    }

    function handleExcelFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(ext)) {
            alert('Invalid file format. Please select an Excel file (.xlsx or .xls) or CSV file.');
            return;
        }

        if (fileDragArea) fileDragArea.style.display = 'none';
        if (loadingIndicator) loadingIndicator.style.display = 'block';

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                let inventorySheetName = null;
                for (const name of workbook.SheetNames) {
                    const nLow = name.toLowerCase();
                    if (nLow === 'inventory' || nLow.includes('du') || nLow.includes('stp') || nLow.includes('stock')) {
                        inventorySheetName = name;
                        break;
                    }
                }

                if (!inventorySheetName) {
                    inventorySheetName = workbook.SheetNames[0];
                }

                const sheet = workbook.Sheets[inventorySheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

                if (!rows || rows.length === 0) {
                    throw new Error("Spreadsheet appears to be empty.");
                }

                let newPressureDus = [];
                let newStps = [];

                // -------------------------------------------------------------
                // Strategy A: Header Row Detection (DU or STP Flat Tables)
                // -------------------------------------------------------------
                let headerRowIdx = -1;
                let isStpTable = false;
                let colMap = {};

                for (let r = 0; r < Math.min(rows.length, 30); r++) {
                    const row = rows[r];
                    if (!row || !Array.isArray(row)) continue;
                    const rowStr = row.map(cell => cell !== null && cell !== undefined ? String(cell).toLowerCase().trim() : '').join(' ');

                    // 1. Check if STP table
                    if ((rowStr.includes('stp') || rowStr.includes('0.75') || rowStr.includes('1.5') || rowStr.includes('submersible')) &&
                        (rowStr.includes('make') || rowStr.includes('sno') || rowStr.includes('type'))) {
                        headerRowIdx = r;
                        isStpTable = true;
                        const makeCols = [];

                        row.forEach((cell, cIdx) => {
                            if (cell === null || cell === undefined) return;
                            const h = String(cell).toLowerCase().trim();
                            if (h.includes('sno') || h.includes('s.no') || h.includes('sl.') || h.includes('sr.')) {
                                colMap.sno = cIdx;
                            } else if (h.includes('type') || h.includes('0.75') || h.includes('1.5') || h.includes('hp') || h.includes('stp')) {
                                colMap.hp_type = cIdx;
                            } else if (h.includes('serial') || h.includes('sr no') || h.includes('sr. no') || h.includes('sl no')) {
                                colMap.serial = cIdx;
                            } else if (h.includes('make') || h.includes('brand')) {
                                makeCols.push(cIdx);
                            } else if (h.includes('model')) {
                                colMap.model = cIdx;
                            } else if (h.includes('date') || h.includes('receipt') || h.includes('receif') || h.includes('dd/mm/yyyy')) {
                                colMap.date = cIdx;
                            } else if (h.includes('new') || h.includes('used') || h.includes('unservice') || h.includes('status') || h.includes('condition')) {
                                colMap.status = cIdx;
                            } else if (h.includes('remark') || h.includes('note') || h.includes('comment')) {
                                colMap.remarks = cIdx;
                            }
                        });

                        // Resolve HPCL template duplicate MAKE header artifact
                        if (colMap.serial === undefined) {
                            if (makeCols.length >= 2) {
                                colMap.make = makeCols[0];
                                colMap.serial = makeCols[1];
                            } else if (makeCols.length === 1 && colMap.model !== undefined && colMap.model - makeCols[0] > 1) {
                                colMap.make = makeCols[0];
                                colMap.serial = makeCols[0] + 1;
                            } else if (makeCols.length === 1) {
                                colMap.make = makeCols[0];
                            }
                        } else {
                            if (makeCols.length > 0) colMap.make = makeCols[0];
                        }
                        break;
                    }

                    // 2. Check if DU table
                    if ((rowStr.includes('serial') || rowStr.includes('variant') || rowStr.includes('dispensing') || rowStr.includes('model')) &&
                        (rowStr.includes('sno') || rowStr.includes('type') || rowStr.includes('make') || rowStr.includes('sr'))) {
                        headerRowIdx = r;
                        isStpTable = false;

                        row.forEach((cell, cIdx) => {
                            if (cell === null || cell === undefined) return;
                            const h = String(cell).toLowerCase().trim();
                            if (h.includes('serial') || h.includes('sr no') || h.includes('sr. no') || h.includes('sl no') || h.includes('sl. no')) {
                                colMap.serial = cIdx;
                            } else if (h.includes('variant') || h.includes('sdm') || h.includes('mpd') || h.includes('hdd') || h.includes('hdm')) {
                                colMap.variant = cIdx;
                            } else if (h.includes('make') || h.includes('brand') || h.includes('manufacturer')) {
                                colMap.make = cIdx;
                            } else if (h.includes('model')) {
                                colMap.model = cIdx;
                            } else if (h.includes('date') || h.includes('receipt') || h.includes('storage point') || h.includes('dd/mm/yyyy')) {
                                colMap.date = cIdx;
                            } else if (h.includes('new') || h.includes('used') || h.includes('unservice') || h.includes('status') || h.includes('condition')) {
                                colMap.status = cIdx;
                            } else if (h.includes('remark') || h.includes('note') || h.includes('comment')) {
                                colMap.remarks = cIdx;
                            } else if (h.includes('type') || h.includes('dispenser') || h.includes('category')) {
                                colMap.type = cIdx;
                            } else if (h.includes('sno') || h.includes('s.no') || h.includes('sr.') || h.includes('sl.')) {
                                colMap.sno = cIdx;
                            }
                        });
                        break;
                    }
                }

                // If dynamic header mapping succeeded and we found at least the serial column
                if (headerRowIdx !== -1 && colMap.serial !== undefined) {
                    for (let r = headerRowIdx + 1; r < rows.length; r++) {
                        const row = rows[r];
                        if (!row || !Array.isArray(row)) continue;

                        const rawSerial = colMap.serial !== undefined && colMap.serial < row.length ? row[colMap.serial] : null;
                        if (!rawSerial) continue;
                        const serial = String(rawSerial).trim();
                        if (!serial || ['nil', 'na', '-', 'none', 'null'].includes(serial.toLowerCase())) continue;

                        const rawSno = colMap.sno !== undefined && colMap.sno < row.length ? row[colMap.sno] : null;
                        const rawMake = colMap.make !== undefined && colMap.make < row.length ? row[colMap.make] : null;
                        const rawModel = colMap.model !== undefined && colMap.model < row.length ? row[colMap.model] : null;
                        const rawDate = colMap.date !== undefined && colMap.date < row.length ? row[colMap.date] : null;
                        const rawStatus = colMap.status !== undefined && colMap.status < row.length ? row[colMap.status] : null;

                        const remarksParts = [];
                        if (colMap.remarks !== undefined && colMap.remarks < row.length) {
                            for (let c = colMap.remarks; c < row.length; c++) {
                                const val = row[c];
                                if (val !== null && val !== undefined && String(val).trim()) {
                                    remarksParts.push(String(val).trim());
                                }
                            }
                        }
                        const remarks = remarksParts.join(' | ');

                        if (isStpTable) {
                            const rawHp = colMap.hp_type !== undefined && colMap.hp_type < row.length ? row[colMap.hp_type] : null;
                            let hpClean = '0.75';
                            if (rawHp) {
                                const hpStr = String(rawHp).trim();
                                if (hpStr.includes('1.5')) hpClean = '1.5';
                                else if (hpStr.includes('0.75')) hpClean = '0.75';
                                else hpClean = hpStr;
                            }

                            newStps.push({
                                sno: rawSno ? Number(rawSno) || (r - headerRowIdx) : (r - headerRowIdx),
                                hp_type: hpClean,
                                make: rawMake ? String(rawMake).trim().toUpperCase() : 'GENERIC',
                                serial_no: serial,
                                model: rawModel ? String(rawModel).trim() : 'NA',
                                receipt_date: parseDateValue(rawDate),
                                status: normalizeStatusValue(rawStatus),
                                remarks: remarks
                            });
                        } else {
                            const rawVariant = colMap.variant !== undefined && colMap.variant < row.length ? row[colMap.variant] : null;
                            const rawType = colMap.type !== undefined && colMap.type < row.length ? row[colMap.type] : null;
                            const typeStr = rawType ? String(rawType).trim() : 'Pressure';

                            newPressureDus.push({
                                sno: rawSno ? Number(rawSno) || (r - headerRowIdx) : (r - headerRowIdx),
                                type: typeStr || 'Pressure',
                                variant: rawVariant ? String(rawVariant).trim() : 'MPD',
                                make: rawMake ? String(rawMake).trim().toUpperCase() : 'GENERIC',
                                model: rawModel ? String(rawModel).trim() : '',
                                serial_no: serial,
                                receipt_date: parseDateValue(rawDate),
                                status: normalizeStatusValue(rawStatus),
                                remarks: remarks
                            });
                        }
                    }
                } else {
                    // -------------------------------------------------------------
                    // Strategy B: Fallback to Multi-section Template Parser
                    // -------------------------------------------------------------
                    let currentSection = null;

                    for (let r = 0; r < rows.length; r++) {
                        const row = rows[r];
                        if (!row || row.length === 0) continue;

                        const rowStr = row.map(cell => cell !== null && cell !== undefined ? String(cell).toLowerCase().trim() : '').join(' ');

                        if (rowStr.includes('suction dispensing units details')) {
                            currentSection = 'suction';
                            r++;
                            continue;
                        } else if (rowStr.includes('pressure type dispensing units details')) {
                            currentSection = 'pressure';
                            r++;
                            continue;
                        } else if (rowStr.includes('stp details')) {
                            currentSection = 'stp';
                            r++;
                            continue;
                        } else if (rowStr.includes('inventory taken by') || rowStr.includes('checked by') || rowStr.includes('approved by')) {
                            currentSection = null;
                            break;
                        }

                        if (!currentSection) continue;

                        const snoCol = row[0] !== null && !isNaN(Number(row[0])) ? 0 : 1;
                        const sno = row[snoCol];
                        if (sno === null || sno === undefined || String(sno).trim() === '' || isNaN(Number(sno))) {
                            continue;
                        }

                        const offset = snoCol === 0 ? -1 : 0;
                        const remarksParts = [];
                        for (let c = 9 + offset; c < row.length; c++) {
                            const cellVal = row[c];
                            if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== '') {
                                remarksParts.push(String(cellVal).trim());
                            }
                        }
                        const remarks = remarksParts.join(' | ');

                        const rawDate = currentSection === 'stp' ? row[6 + offset] : row[7 + offset];
                        const receiptDate = parseDateValue(rawDate);

                        if (currentSection === 'pressure') {
                            const rawSer = row[6 + offset];
                            if (!rawSer) continue;
                            newPressureDus.push({
                                sno: Number(sno),
                                type: 'Pressure',
                                variant: row[3 + offset] ? String(row[3 + offset]).trim() : 'Dispenser',
                                make: row[4 + offset] ? String(row[4 + offset]).trim().toUpperCase() : 'UNKNOWN',
                                model: row[5 + offset] ? String(row[5 + offset]).trim() : '',
                                serial_no: String(rawSer).trim(),
                                receipt_date: receiptDate,
                                status: normalizeStatusValue(row[8 + offset]),
                                remarks: remarks
                            });
                        } else if (currentSection === 'stp') {
                            const rawSer = row[4 + offset];
                            if (!rawSer) continue;
                            newStps.push({
                                sno: Number(sno),
                                hp_type: row[2 + offset] ? String(row[2 + offset]).trim() : '0.75',
                                make: row[3 + offset] ? String(row[3 + offset]).trim().toUpperCase() : 'UNKNOWN',
                                serial_no: String(rawSer).trim(),
                                model: row[5 + offset] ? String(row[5 + offset]).trim() : '',
                                receipt_date: receiptDate,
                                status: normalizeStatusValue(row[8 + offset]),
                                remarks: remarks
                            });
                        }
                    }
                }

                // -------------------------------------------------------------
                // Update or Append into Inventory Database
                // -------------------------------------------------------------
                if (newPressureDus.length > 0 || newStps.length > 0) {
                    let dusAdded = 0;
                    let dusUpdated = 0;
                    let stpsAdded = 0;
                    let stpsUpdated = 0;

                    if (!db.inventory) db.inventory = {};
                    if (!Array.isArray(db.inventory.pressure_dus)) db.inventory.pressure_dus = [];
                    if (!Array.isArray(db.inventory.stps)) db.inventory.stps = [];
                    if (!Array.isArray(db.inword)) db.inword = [];

                    // Process DUs: Update existing or append new
                    newPressureDus.forEach(item => {
                        const sClean = safeLower(item.serial_no);
                        if (!sClean) return;

                        const existingIdx = db.inventory.pressure_dus.findIndex(x => safeLower(x.serial_no) === sClean);

                        if (existingIdx !== -1) {
                            const target = db.inventory.pressure_dus[existingIdx];
                            if (item.variant) target.variant = item.variant;
                            if (item.make && item.make !== 'GENERIC') target.make = item.make;
                            if (item.model) target.model = item.model;
                            if (item.status) target.status = item.status;
                            if (item.remarks) target.remarks = item.remarks;
                            if (item.receipt_date) target.receipt_date = item.receipt_date;
                            if (item.type) target.type = item.type;
                            dusUpdated++;
                        } else {
                            db.inventory.pressure_dus.push(item);
                            dusAdded++;
                            db.inword.push({
                                item: `${item.variant || 'MPD'} Dispenser (${item.make})`,
                                received_from: `Excel Import (${item.make})`,
                                received_date: item.receipt_date || todayLocal(),
                                qty: 1
                            });
                        }
                    });

                    // Process STPs: Update existing or append new
                    newStps.forEach(item => {
                        const sClean = safeLower(item.serial_no);
                        if (!sClean) return;

                        const existingIdx = db.inventory.stps.findIndex(x => safeLower(x.serial_no) === sClean);

                        if (existingIdx !== -1) {
                            const target = db.inventory.stps[existingIdx];
                            if (item.hp_type) target.hp_type = item.hp_type;
                            if (item.make && item.make !== 'GENERIC') target.make = item.make;
                            if (item.model) target.model = item.model;
                            if (item.status) target.status = item.status;
                            if (item.remarks) target.remarks = item.remarks;
                            if (item.receipt_date) target.receipt_date = item.receipt_date;
                            stpsUpdated++;
                        } else {
                            db.inventory.stps.push(item);
                            stpsAdded++;
                            db.inword.push({
                                item: `STP Pump ${item.hp_type} HP (${item.make})`,
                                received_from: `Excel Import (${item.make})`,
                                received_date: item.receipt_date || todayLocal(),
                                qty: 1
                            });
                        }
                    });

                    // Re-sequence sequential numbers
                    db.inventory.pressure_dus.forEach((d, idx) => d.sno = idx + 1);
                    db.inventory.stps.forEach((s, idx) => s.sno = idx + 1);

                    saveState();
                    refreshAllViews();

                    let summaryMsg = "Inventory Updated Successfully!\n";
                    if (stpsUpdated > 0) summaryMsg += `\n• Existing STPs Updated: ${stpsUpdated}`;
                    if (stpsAdded > 0) summaryMsg += `\n• New STPs Added: ${stpsAdded}`;
                    if (dusUpdated > 0) summaryMsg += `\n• Existing DUs Updated: ${dusUpdated}`;
                    if (dusAdded > 0) summaryMsg += `\n• New DUs Added: ${dusAdded}`;
                    summaryMsg += `\n• Active STPs in Depot: ${db.inventory.stps.length}`;
                    summaryMsg += `\n• Active DUs in Depot: ${db.inventory.pressure_dus.length}`;

                    const toastDetail = isStpTable
                        ? `${stpsUpdated} STPs updated, ${stpsAdded} added.`
                        : `${dusUpdated} DUs updated, ${dusAdded} added.`;
                    showToast(`Import Success: ${toastDetail}`);
                    alert(summaryMsg);

                    if (importModal) importModal.classList.remove('active');
                } else {
                    alert(`No valid dispensing units or STPs parsed from '${file.name}'. Please ensure your file has columns like 'TYPE OF STPs', 'SERIAL NO', 'MAKE', 'STATUS', etc.`);
                    if (fileDragArea) fileDragArea.style.display = 'block';
                    if (loadingIndicator) loadingIndicator.style.display = 'none';
                }

            } catch (err) {
                console.error("Excel import error:", err);
                alert("Error parsing spreadsheet: " + err.message);
                if (fileDragArea) fileDragArea.style.display = 'block';
                if (loadingIndicator) loadingIndicator.style.display = 'none';
            }
        };

        reader.onerror = () => {
            alert("Error reading file.");
            if (fileDragArea) fileDragArea.style.display = 'block';
            if (loadingIndicator) loadingIndicator.style.display = 'none';
        };

        reader.readAsArrayBuffer(file);
    }
}

// ==========================================================================
// Backup & Google Sheets Sync Controls
// ==========================================================================
function setupBackendControls() {
    const sheetIdInput = document.getElementById('sheets-spreadsheet-id');
    const autoSyncCheckbox = document.getElementById('sheets-auto-sync');
    const scanDirInput = document.getElementById('scan-directory-path');
    const mergeModeCheckbox = document.getElementById('scan-merge-mode');
    
    if (!sheetIdInput) return;

    const btnSaveConfig = () => {
        const url = sheetIdInput.value.trim();
        const autoSync = autoSyncCheckbox ? autoSyncCheckbox.checked : true;
        
        localStorage.setItem('hpcl_sheets_url', url);
        localStorage.setItem('hpcl_sheets_auto_sync', autoSync ? 'true' : 'false');
        
        showToast("Sync settings saved!");
        updateSheetsStatusBadge();
        
        if (url && autoSync && navigator.onLine) {
            autoPushGoogleSheets();
        }

        if (backendActive) {
            fetch('/api/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    sheet_id: url,
                    auto_sync: autoSync,
                    scan_directory: scanDirInput ? scanDirInput.value.trim() : '',
                    merge_mode: mergeModeCheckbox ? mergeModeCheckbox.checked : true
                })
            }).catch(() => {});
        }
    };

    [sheetIdInput, scanDirInput].forEach(elem => {
        if (elem) elem.addEventListener('blur', btnSaveConfig);
    });
    sheetIdInput.addEventListener('change', btnSaveConfig);
    [autoSyncCheckbox, mergeModeCheckbox].forEach(elem => {
        if (elem) elem.addEventListener('change', btnSaveConfig);
    });

    const persistUrlOnly = (val) => {
        localStorage.setItem('hpcl_sheets_url', val);
        updateSheetsStatusBadge();
    };

    const btnPush = document.getElementById('btn-push-sheets');
    if (btnPush) {
        btnPush.addEventListener('click', async () => {
            const url = sheetIdInput.value.trim();
            if (!url) {
                alert("Please paste your Google Apps Script /exec Web App URL first.");
                return;
            }
            persistUrlOnly(url);

            btnPush.disabled = true;
            const originalText = btnPush.innerHTML;
            btnPush.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Pushing...';

            try {
                await pushToGoogleSheets(url, db);
                showToast("Pushed local data to Google Sheets!");
            } catch (err) {
                alert("Push to Sheets error:\n" + err.message);
            } finally {
                btnPush.disabled = false;
                btnPush.innerHTML = originalText;
            }
        });
    }

    const btnTestSheets = document.getElementById('btn-test-sheets');
    if (btnTestSheets) {
        btnTestSheets.addEventListener('click', async () => {
            const url = sheetIdInput.value.trim();
            if (!url) {
                alert("Please enter a Google Sheet URL or Apps Script /exec Web App link first.");
                return;
            }
            persistUrlOnly(url);
            btnTestSheets.disabled = true;
            const originalHtml = btnTestSheets.innerHTML;
            btnTestSheets.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Testing...';
            const startTime = Date.now();
            try {
                if (isAppsScriptUrl(url)) {
                    const testUrl = url + (url.includes('?') ? '&' : '?') + 'action=pull';
                    const res = await fetch(testUrl, { method: 'GET', redirect: 'follow' });
                    const txt = await res.text();
                    let parsed = null;
                    try { parsed = JSON.parse(txt); } catch (e) {}
                    const pingTime = Date.now() - startTime;
                    if (!res.ok || !parsed || parsed.status === 'error') {
                        throw new Error((parsed && parsed.message) || `Invalid response from Apps Script (HTTP ${res.status}). Check 'Who has access: Anyone' in Deploy -> Manage deployments.`);
                    }
                    showToast(`Apps Script connected (${pingTime}ms)`);
                    alert(`Connection Test Passed!\nApps Script /exec Web App responded with valid JSON in ${pingTime}ms.\nBoth Push and Pull are ready.`);
                } else {
                    await fetch(url, { method: 'HEAD', mode: 'no-cors' });
                    const pingTime = Date.now() - startTime;
                    showToast(`Sheet URL reachable (${pingTime}ms) — pull only`);
                    alert(`Spreadsheet Link Reachable (${pingTime}ms) — Pull Only.\n\nNote: A google.com/spreadsheets link can only Pull data. To enable 'Push now' and Drive photo uploads, deploy google_apps_script.js as a Web App and paste the /exec URL.`);
                }
            } catch (e) {
                alert(`Connection Test Failed:\n${e.message || e}\n\nMake sure the Web App is deployed with 'Execute as: Me' and 'Who has access: Anyone'.`);
            } finally {
                btnTestSheets.disabled = false;
                btnTestSheets.innerHTML = originalHtml;
            }
        });
    }

    const btnCopyScript = document.getElementById('btn-copy-script');
    if (btnCopyScript) {
        btnCopyScript.addEventListener('click', async () => {
            try {
                const res = await fetch('./google_apps_script.js');
                if (!res.ok) throw new Error(`HTTP ${res.status}`);
                const code = await res.text();
                copyToClipboard(code, 'Google Apps Script v3.0');
            } catch (err) {
                alert('Could not load google_apps_script.js automatically. You can copy it directly from google_apps_script.js in the repository.');
            }
        });
    }

    const btnViewScript = document.getElementById('btn-view-script');
    if (btnViewScript) {
        btnViewScript.addEventListener('click', () => {
            const scriptInstructions = `Google Apps Script v3.0 Setup (Enables Push + Pull + Photo Upload):

1. Open your Google Sheet -> Extensions -> Apps Script.
2. Delete existing code, click "Copy Apps Script" in this app, and paste it in.
3. Click Deploy -> New deployment (or Manage deployments -> Edit -> New version).
4. Type: Web app
5. Execute as: Me
6. Who has access: Anyone  (Required for Push & Pull from any device)
7. Click Deploy, grant permissions (including Google Drive for photos), and copy the URL ending in /exec.
8. Paste that /exec URL into the "Apps Script Web App URL" box and click "Push now".`;
            alert(scriptInstructions);
        });
    }

    const btnPull = document.getElementById('btn-pull-sheets');
    if (btnPull) {
        btnPull.addEventListener('click', async () => {
            const sheetUrl = sheetIdInput.value.trim();
            if (!sheetUrl) {
                alert("Please paste your Google Sheet URL or Apps Script Link first.");
                return;
            }
            persistUrlOnly(sheetUrl);
            btnPull.disabled = true;
            const originalText = btnPull.innerHTML;
            btnPull.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Pulling...';
            
            try {
                const pulledData = await pullFromGoogleSheets(sheetUrl);
                if (pulledData) {
                    const inv = pulledData.inventory || {};
                    const pDus = (inv.pressure_dus || pulledData.pressure_dus || []).length;
                    const pStps = (inv.stps || pulledData.stps || []).length;
                    const pEmrs = (pulledData.gate_passes || []).length;
                    const pLetters = (pulledData.dispatch_letters || []).length;
                    if (!confirm(`Google Sheets returned ${pDus} DUs, ${pStps} STPs, ${pEmrs} EMRs and ${pLetters} dispatch letters.\n\nReplace the data on this device with it? (Tip: export a JSON backup first if unsure.)`)) {
                        return;
                    }
                    if (!pulledData.inventory) {
                        pulledData.inventory = { pressure_dus: pulledData.pressure_dus || [], stps: pulledData.stps || [] };
                        delete pulledData.pressure_dus;
                        delete pulledData.stps;
                    }
                    db = pulledData;
                    standardizeDatabase();
                    saveState();
                    refreshAllViews();
                    showToast("Database pulled from Google Sheets successfully!");
                    alert("Pull Successful! Inventory database updated with latest data from Google Sheets.");
                } else {
                    throw new Error("No valid data received from Google Sheets.");
                }
            } catch (error) {
                console.error("Pull error:", error);
                alert("Pull Error: " + error.message);
            } finally {
                btnPull.disabled = false;
                btnPull.innerHTML = originalText;
            }
        });
    }

    const btnScan = document.getElementById('btn-scan-directory');
    if (btnScan) {
        btnScan.addEventListener('click', async () => {
            if (!backendActive) {
                alert("Backend server is not running for directory scan.");
                return;
            }
            const scanDir = scanDirInput.value.trim();
            if (!scanDir) {
                alert("Please specify a local folder path to scan.");
                return;
            }
            
            btnScan.disabled = true;
            const originalText = btnScan.innerHTML;
            btnScan.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Scanning...';
            
            try {
                const response = await fetch('/api/scan-directory', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        directory_path: scanDir,
                        merge_mode: mergeModeCheckbox ? mergeModeCheckbox.checked : true
                    })
                });
                
                if (response.ok) {
                    const result = await response.json();
                    if (result.files_found === 0) {
                        alert("Scan completed. No Excel spreadsheets (.xlsx or .xls) found in that directory.");
                    } else {
                        const dataRes = await fetch('/api/data');
                        if (dataRes.ok) {
                            db = await dataRes.json();
                            standardizeDatabase();
                            saveState();
                            refreshAllViews();
                        }
                        alert(`Scanned directory successfully! Found ${result.files_found} files.`);
                    }
                } else {
                    const err = await response.json();
                    alert("Scan failed: " + (err.detail || response.statusText));
                }
            } catch (error) {
                alert("Scan request failed: " + error.message);
            } finally {
                btnScan.disabled = false;
                btnScan.innerHTML = originalText;
            }
        });
    }
}

function updateSheetsStatusBadge(isSyncing = false, customStatus = null) {
    const badge = document.getElementById('sheets-connection-status');
    const textElem = document.getElementById('sheets-instructions-text');
    const topPillText = document.getElementById('top-sync-text');
    const topPillDot = document.querySelector('#top-sync-pill .sync-dot');
    if (!badge) return;

    const url = localStorage.getItem('hpcl_sheets_url');
    const isOnline = navigator.onLine;

    if (customStatus) {
        badge.textContent = customStatus;
        badge.className = "status-badge status-warning";
        if (topPillText) topPillText.textContent = customStatus;
        if (topPillDot) topPillDot.style.backgroundColor = "var(--color-warning)";
        return;
    }

    if (!isOnline) {
        badge.textContent = "Offline";
        badge.className = "status-badge status-warning";
        badge.style.backgroundColor = "var(--color-warning-bg)";
        badge.style.color = "var(--color-warning)";
        if (topPillText) topPillText.textContent = "Offline, saving on this device";
        if (topPillDot) topPillDot.style.backgroundColor = "var(--color-warning)";
        if (textElem) {
            textElem.innerHTML = `<strong>You are offline.</strong> Everything is saved on this device and will be pushed to Google Sheets when the connection is back.`;
        }
    } else if (url && !isAppsScriptUrl(url)) {
        badge.textContent = "Pull only";
        badge.className = "status-badge status-info";
        badge.style.backgroundColor = "rgba(59, 130, 246, 0.15)";
        badge.style.color = "#60a5fa";
        if (topPillText) topPillText.textContent = "Sheet linked (pull only)";
        if (topPillDot) topPillDot.style.backgroundColor = "var(--color-info)";
        if (textElem) {
            textElem.innerHTML = `<strong>Spreadsheet link saved.</strong> You can pull data from it. To push changes and upload dispatch photos, deploy <code>google_apps_script.js</code> as a Web App and paste its <code>/exec</code> URL here.`;
        }
    } else if (url) {
        badge.textContent = "Connected";
        badge.className = "status-badge status-success";
        badge.style.backgroundColor = "var(--color-success-bg)";
        badge.style.color = "var(--color-success)";
        if (topPillText) topPillText.textContent = "Syncing to Google Sheets";
        if (topPillDot) topPillDot.style.backgroundColor = "var(--color-success)";
        if (textElem) {
            textElem.innerHTML = `<strong>Connected.</strong> Changes, dispatch letters and EMRs are pushed to your Google Sheet; dispatch photos are uploaded to Google Drive.`;
        }
    } else {
        badge.textContent = "Not linked";
        badge.className = "status-badge status-info";
        badge.style.backgroundColor = "rgba(59, 130, 246, 0.15)";
        badge.style.color = "#60a5fa";
        if (topPillText) topPillText.textContent = "Saved on this device only";
        if (topPillDot) topPillDot.style.backgroundColor = "var(--color-info)";
        if (textElem) {
            textElem.innerHTML = `<strong>Not linked.</strong> Data stays in this browser only. Paste the Apps Script Web App URL above so the office and the depot share one register.`;
        }
    }
}

// ==========================================================================
// EMR (Equipment Movement Record) Management Module (Multi-Asset Capable)
// ==========================================================================
let currentEmrAssets = [];

// Snapshot of an in-store asset as it leaves on an EMR. Condition, receipt date and remarks are
// kept so that cancelling the EMR can put the asset back exactly as it was.
function buildEmrAsset(category, found) {
    const hp = safeStr(found.hp_type || found.variant || '0.75').replace(/\s*HP$/i, '');
    return {
        category: category,
        serial_no: found.serial_no,
        make: found.make || '',
        model: found.model || '',
        variant: category === 'DU' ? (found.variant || 'MPD') : `${hp} HP`,
        hp_type: category === 'DU' ? '' : hp,
        type: category === 'DU' ? (found.type || 'Pressure') : '',
        receipt_date: found.receipt_date || '',
        status: found.status || 'N',
        remarks: found.remarks || ''
    };
}

function renderEmrSelectedAssetsTable() {
    const tbody = document.getElementById('gp-selected-assets-tbody');
    const container = document.getElementById('gp-selected-assets-container');
    const badge = document.getElementById('gp-selected-assets-badge');
    const countNum = document.getElementById('gp-asset-count-num');
    if (!tbody || !container) return;

    tbody.innerHTML = '';
    const count = currentEmrAssets.length;

    if (badge) {
        badge.textContent = `${count} asset${count === 1 ? '' : 's'}`;
        badge.className = count > 0 ? "badge badge-success" : "badge badge-info";
    }
    if (countNum) countNum.textContent = count;

    if (count === 0) {
        container.style.display = 'none';
        return;
    }

    container.style.display = 'block';

    currentEmrAssets.map(escapeRecord).forEach((asset, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="text-muted">${idx + 1}</td>
            <td>${asset.category === 'DU' ? 'DU' : 'STP'}</td>
            <td><span class="serial">${asset.serial_no}</span></td>
            <td>${asset.make || ''} ${asset.model || ''} <span class="text-muted">${asset.variant || asset.hp_type || ''}</span></td>
            <td class="text-right"><button type="button" class="btn-icon btn-delete gp-remove-asset-btn" data-serial="${asset.serial_no}" title="Remove from EMR" aria-label="Remove ${asset.serial_no}"><i class="fa-solid fa-xmark"></i></button></td>
        `;
        tbody.appendChild(tr);
    });
}

function addAssetToEmr(category, serialNo) {
    if (!serialNo) {
        alert("Please select a Serial Number from the dropdown first.");
        return;
    }

    const exists = currentEmrAssets.some(a => safeLower(a.serial_no) === safeLower(serialNo));
    if (exists) {
        alert(`Serial Number '${serialNo}' is already added to this EMR.`);
        return;
    }

    let found = null;
    if (category === 'DU') {
        found = (db.inventory.pressure_dus || []).find(i => safeLower(i.serial_no) === safeLower(serialNo));
    } else {
        found = (db.inventory.stps || []).find(i => safeLower(i.serial_no) === safeLower(serialNo));
    }

    if (!found) {
        alert(`Serial Number '${serialNo}' not found in active Store Inventory.`);
        return;
    }

    currentEmrAssets.push(buildEmrAsset(category, found));

    const serSelect = document.getElementById('gp-serial-select');
    if (serSelect) serSelect.value = '';

    const previewCard = document.getElementById('gp-item-preview-card');
    if (previewCard) previewCard.style.display = 'none';

    renderEmrSelectedAssetsTable();
    populateGatePassSerialDropdown(category);
}

function removeAssetFromEmr(serialNo) {
    currentEmrAssets = currentEmrAssets.filter(a => safeLower(a.serial_no) !== safeLower(serialNo));
    renderEmrSelectedAssetsTable();
    const catSelect = document.getElementById('gp-category');
    if (catSelect) populateGatePassSerialDropdown(catSelect.value);
}

function renderGatePassView() {
    if (!db.gate_passes) db.gate_passes = [];
    
    const totalGPs = db.gate_passes.length;
    const duDispatched = db.gate_passes.reduce((sum, g) => sum + (g.items ? g.items.filter(i => i.category === 'DU').length : (g.category === 'DU' ? 1 : 0)), 0);
    const stpDispatched = db.gate_passes.reduce((sum, g) => sum + (g.items ? g.items.filter(i => i.category === 'STP').length : (g.category === 'STP' ? 1 : 0)), 0);
    
    const statTotal = document.getElementById('gp-stat-total');
    const statDus = document.getElementById('gp-stat-dus');
    const statStps = document.getElementById('gp-stat-stps');
    if (statTotal) statTotal.textContent = totalGPs;
    if (statDus) statDus.textContent = duDispatched;
    if (statStps) statStps.textContent = stpDispatched;
    const statPhotos = document.getElementById('gp-stat-photos');
    if (statPhotos) statPhotos.textContent = db.gate_passes.filter(g => Array.isArray(g.photos) && g.photos.length).length;
    
    const latestElem = document.getElementById('gp-stat-latest');
    const latestSubElem = document.getElementById('gp-stat-latest-sub');
    if (latestElem && latestSubElem) {
        if (totalGPs > 0) {
            const latest = db.gate_passes[totalGPs - 1];
            latestElem.textContent = safeStr(latest.emr_no || latest.gate_pass_no) || '-';
            const assetCount = (latest.items && latest.items.length) ? latest.items.length : 1;
            latestSubElem.textContent = `${assetCount} Asset${assetCount > 1 ? 's' : ''} (${formatDisplayDate(latest.out_date) || 'date n/a'})`;
        } else {
            latestElem.textContent = '-';
            latestSubElem.textContent = 'No dispatch yet';
        }
    }

    const vendorSelect = document.getElementById('gp-filter-vendor');
    if (vendorSelect) {
        const currentVendor = vendorSelect.value;
        const vendors = new Set();
        db.gate_passes.forEach(g => {
            if (safeStr(g.vendor)) vendors.add(safeStr(g.vendor));
        });
        vendorSelect.innerHTML = '<option value="all">All Vendors</option>';
        Array.from(vendors).sort().forEach(v => {
            const opt = document.createElement('option');
            opt.value = v;
            opt.textContent = v;
            vendorSelect.appendChild(opt);
        });
        if (currentVendor && vendors.has(currentVendor)) {
            vendorSelect.value = currentVendor;
        }
    }

    const tbody = document.getElementById('gate-pass-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const categoryFilter = document.getElementById('gp-filter-category') ? document.getElementById('gp-filter-category').value : 'all';
    const vendorFilter = document.getElementById('gp-filter-vendor') ? document.getElementById('gp-filter-vendor').value : 'all';
    const searchVal = document.getElementById('gp-search') ? document.getElementById('gp-search').value.toLowerCase().trim() : '';

    let list = [...db.gate_passes];

    if (categoryFilter !== 'all') {
        list = list.filter(g => {
            if (g.items && g.items.length > 0) {
                return g.items.some(i => i.category === categoryFilter);
            }
            return g.category === categoryFilter;
        });
    }
    if (vendorFilter !== 'all') {
        list = list.filter(g => safeStr(g.vendor) === vendorFilter);
    }
    const photoFilterEl = document.getElementById('gp-filter-photo');
    const photoFilter = photoFilterEl ? photoFilterEl.value : 'all';
    if (photoFilter === 'missing') list = list.filter(g => !(Array.isArray(g.photos) && g.photos.length));
    if (photoFilter === 'has') list = list.filter(g => Array.isArray(g.photos) && g.photos.length);
    if (searchVal !== '') {
        list = list.filter(g => {
            const allSerials = safeLower(g.items ? g.items.map(i => i.serial_no).join(' ') : g.serial_no);
            return safeLower(g.emr_no).includes(searchVal) ||
                   safeLower(g.gate_pass_no).includes(searchVal) ||
                   allSerials.includes(searchVal) ||
                   safeLower(g.destination).includes(searchVal) ||
                   safeLower(g.vendor).includes(searchVal) ||
                   safeLower(g.make).includes(searchVal) ||
                   safeLower(g.model).includes(searchVal) ||
                   safeLower(g.vehicle_no).includes(searchVal) ||
                   safeLower(g.dl_no).includes(searchVal);
        });
    }

    // Sort EMRs
    list.sort((a, b) => {
        let valA = a[gpSortColumn] || (gpSortColumn === 'emr_no' ? a.gate_pass_no : '');
        let valB = b[gpSortColumn] || (gpSortColumn === 'emr_no' ? b.gate_pass_no : '');

        if (gpSortColumn === 'out_date') {
            return gpSortDirection === 'asc'
                ? safeDateTimestamp(valA) - safeDateTimestamp(valB)
                : safeDateTimestamp(valB) - safeDateTimestamp(valA);
        }
        valA = safeLower(valA);
        valB = safeLower(valB);
        return gpSortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
    });

    if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="empty-row"><strong>No EMRs match</strong>EMRs are created when the depot records a dispatch against a letter, or with Issue EMR without letter.</td></tr>`;
        return;
    }

    list.forEach(rawItem => {
        const tr = document.createElement('tr');
        const item = escapeRecord(rawItem);
        const itemsList = (rawItem.items && rawItem.items.length > 0) ? rawItem.items.map(escapeRecord) : [{
            category: item.category,
            serial_no: item.serial_no,
            make: item.make,
            model: item.model,
            variant: item.variant
        }];

        const duCount = itemsList.filter(i => safeStr(i.category).includes('DU')).length;
        const stpCount = itemsList.filter(i => safeStr(i.category).includes('STP')).length;
        const catParts = [];
        if (duCount) catParts.push(`${duCount} DU`);
        if (stpCount) catParts.push(`${stpCount} STP`);
        if (Array.isArray(rawItem.other_items) && rawItem.other_items.length) catParts.push(`${rawItem.other_items.length} other`);
        const catBadge = catParts.join(' + ') || escapeHtml(rawItem.category) || '—';

        const serialsHtml = itemsList.filter(it => it.serial_no).map(it => `<span class="serial-chip">${it.serial_no}</span>`).join('') || '<span class="text-muted">—</span>';
        const makeModelHtml = itemsList.map(it => `${it.make || ''} ${it.model || ''}`.trim()).filter(Boolean).filter((v, i, a) => a.indexOf(v) === i).join(', ') || item.make || '—';

        const emrDisplay = item.emr_no || item.gate_pass_no || '-';
        const gpDisplay = item.gate_pass_no && item.gate_pass_no !== item.emr_no ? item.gate_pass_no : '';
        const photoCount = Array.isArray(rawItem.photos) ? rawItem.photos.length : 0;
        const letterLine = item.dl_no ? `<div style="font-size:.82rem;"><a href="#letters" class="goto-letter" data-dl="${item.dl_no}">${item.dl_no}</a></div>` : '';

        tr.innerHTML = `
            <td style="white-space:nowrap;">
                <span class="copyable-serial" data-copy="${emrDisplay}" title="Copy EMR no">${emrDisplay}<i class="fa-regular fa-copy"></i></span>
                ${letterLine}
            </td>
            <td style="white-space:nowrap;">${formatDisplayDate(item.out_date) || '—'}</td>
            <td>${catBadge}</td>
            <td style="max-width:280px;">${serialsHtml}</td>
            <td style="font-size:.88rem;">${makeModelHtml}</td>
            <td>${item.vendor || '—'}</td>
            <td>${item.destination || '—'}</td>
            <td class="gp-vehicle" style="font-size:.86rem; color:var(--ink-2); white-space:nowrap;">${gpDisplay ? `<div>GP ${gpDisplay}</div>` : ''}<div>${item.vehicle_no || '<span class="text-muted">No vehicle</span>'}</div></td>
            <td>
                <div class="actions-cell">
                    <button class="btn-icon gp-photos-btn" data-emr="${emrDisplay}" title="${photoCount ? 'View dispatch photos' : 'Add dispatch photo'}" aria-label="Photos for ${emrDisplay}" ${photoCount ? '' : 'style="border-color:var(--yellow); color:var(--yellow-ink);"'}>
                        <i class="fa-solid fa-camera"></i>${photoCount ? `<span class="count">${photoCount}</span>` : ''}
                    </button>
                    <button class="btn-icon gp-print-btn" title="Print EMR voucher" data-emr="${emrDisplay}" aria-label="Print ${emrDisplay}"><i class="fa-solid fa-print"></i></button>
                    <button class="btn-icon btn-delete gp-revert-btn" title="Cancel EMR and return assets to stock" data-emr="${emrDisplay}" aria-label="Cancel ${emrDisplay}"><i class="fa-solid fa-rotate-left"></i></button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const gpTable = document.getElementById('gate-pass-table');
    if (gpTable && !gpTable.dataset.copyListener) {
        gpTable.dataset.copyListener = 'true';
        gpTable.addEventListener('click', (e) => {
            const copyEl = e.target.closest('.copyable-serial');
            if (copyEl) {
                copyToClipboard(copyEl.getAttribute('data-copy'), 'EMR Number');
                return;
            }
            const printBtn = e.target.closest('.gp-print-btn');
            if (printBtn) {
                openGatePassPrintModal(printBtn.getAttribute('data-emr'));
                return;
            }
            const revertBtn = e.target.closest('.gp-revert-btn');
            if (revertBtn) {
                revertGatePass(revertBtn.getAttribute('data-emr'));
                return;
            }
            const photosBtn = e.target.closest('.gp-photos-btn');
            if (photosBtn && typeof openPhotosModal === 'function') {
                openPhotosModal(photosBtn.getAttribute('data-emr'));
                return;
            }
            const letterLink = e.target.closest('.goto-letter');
            if (letterLink && typeof gotoLetter === 'function') {
                e.preventDefault();
                gotoLetter(letterLink.getAttribute('data-dl'));
            }
        });
    }
}

function setupGatePassHandlers() {
    const createBtn = document.getElementById('create-gate-pass-btn');
    const modal = document.getElementById('gate-pass-modal');
    const closeBtn = document.getElementById('gate-pass-modal-close');
    const cancelBtn = document.getElementById('gp-modal-cancel');
    const form = document.getElementById('gate-pass-form');
    
    const categorySelect = document.getElementById('gp-category');
    const serialSelect = document.getElementById('gp-serial-select');
    const addAssetBtn = document.getElementById('gp-add-asset-btn');
    const selectedTbody = document.getElementById('gp-selected-assets-tbody');
    
    if (createBtn) {
        createBtn.addEventListener('click', () => {
            openGatePassModal();
        });
    }

    if (closeBtn && modal) {
        closeBtn.addEventListener('click', () => {
            modal.classList.remove('active');
        });
    }

    if (cancelBtn && modal) {
        cancelBtn.addEventListener('click', () => {
            modal.classList.remove('active');
        });
    }

    if (categorySelect) {
        categorySelect.addEventListener('change', () => {
            populateGatePassSerialDropdown(categorySelect.value);
        });
    }

    if (serialSelect) {
        serialSelect.addEventListener('change', () => {
            updateGatePassPreviewCard();
        });
    }

    if (addAssetBtn) {
        addAssetBtn.addEventListener('click', () => {
            const cat = categorySelect ? categorySelect.value : 'DU';
            const ser = serialSelect ? serialSelect.value : '';
            addAssetToEmr(cat, ser);
        });
    }

    if (selectedTbody) {
        selectedTbody.addEventListener('click', (e) => {
            const removeBtn = e.target.closest('.gp-remove-asset-btn');
            if (removeBtn) {
                const serial = removeBtn.getAttribute('data-serial');
                if (serial) removeAssetFromEmr(serial);
            }
        });
    }

    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            handleGatePassSubmit();
        });
    }

    const gpSearch = document.getElementById('gp-search');
    const gpFilterCat = document.getElementById('gp-filter-category');
    const gpFilterVend = document.getElementById('gp-filter-vendor');

    if (gpSearch) gpSearch.addEventListener('input', () => renderGatePassView());
    if (gpFilterCat) gpFilterCat.addEventListener('change', () => renderGatePassView());
    if (gpFilterVend) gpFilterVend.addEventListener('change', () => renderGatePassView());
    const gpFilterPhoto = document.getElementById('gp-filter-photo');
    if (gpFilterPhoto) gpFilterPhoto.addEventListener('change', () => renderGatePassView());

    const printClose = document.getElementById('gp-print-modal-close');
    const printDoBtn = document.getElementById('do-print-gp-btn');
    const printModal = document.getElementById('gate-pass-print-modal');

    if (printClose && printModal) {
        printClose.addEventListener('click', () => printModal.classList.remove('active'));
    }
    if (printDoBtn) {
        printDoBtn.addEventListener('click', () => {
            window.print();
        });
    }
}

function openGatePassModal(category = 'DU', preselectSerial = null) {
    const modal = document.getElementById('gate-pass-modal');
    if (!modal) return;

    currentEmrAssets = [];

    const catSelect = document.getElementById('gp-category');
    if (catSelect) catSelect.value = category;

    // If preselectSerial is given from "Issue EMR" button, add directly to assets list
    if (preselectSerial) {
        let found = null;
        if (category === 'DU') {
            found = (db.inventory.pressure_dus || []).find(i => safeLower(i.serial_no) === safeLower(preselectSerial));
        } else {
            found = (db.inventory.stps || []).find(i => safeLower(i.serial_no) === safeLower(preselectSerial));
        }
        if (found) {
            currentEmrAssets.push(buildEmrAsset(category, found));
        }
    }

    populateGatePassSerialDropdown(category);
    renderEmrSelectedAssetsTable();

    // Auto-generate strictly unique EMR Number
    const todayStr = todayIso().replace(/-/g, '');
    // Next running number for today, e.g. EMR-20261127-002
    const prefix = `EMR-${todayStr}-`;
    let seq = 1;
    (db.gate_passes || []).forEach(g => {
        const n = safeStr(g.emr_no);
        if (n.startsWith(prefix)) seq = Math.max(seq, (parseInt(n.slice(prefix.length), 10) || 0) + 1);
    });
    const autoNumber = `${prefix}${String(seq).padStart(3, '0')}`;

    const numInput = document.getElementById('gp-number');
    if (numInput) numInput.value = autoNumber;

    const outDateInput = document.getElementById('gp-out-date');
    if (outDateInput) outDateInput.value = todayIso();

    const gpEmrInput = document.getElementById('gp-emr-no');
    if (gpEmrInput) gpEmrInput.value = '';
    ['gp-vendor', 'gp-destination', 'gp-vehicle-no', 'gp-driver', 'gp-remarks'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    if (typeof resetPhotoCapture === 'function') resetPhotoCapture('gp');

    modal.classList.add('active');
}

function quickIssueGatePass(category, serialNo) {
    const navItem = document.querySelector('.nav-item[data-tab="emr"]') || document.querySelector('.nav-item[data-tab="gate-pass"]');
    if (navItem) navItem.click();
    openGatePassModal(category, serialNo);
}

function populateGatePassSerialDropdown(category, preselectSerial = null) {
    const serialSelect = document.getElementById('gp-serial-select');
    if (!serialSelect) return;
    serialSelect.innerHTML = '<option value="">-- Choose Available Item --</option>';

    let availableItems = [];
    if (category === 'DU') {
        availableItems = (db.inventory.pressure_dus || []).map(item => ({
            serial_no: item.serial_no,
            label: `${item.serial_no} | ${item.make} - ${item.model} (${item.variant})`,
            item: item
        }));
    } else {
        availableItems = (db.inventory.stps || []).map(item => ({
            serial_no: item.serial_no,
            label: `${item.serial_no} | ${item.make} - ${item.model} (${item.hp_type} HP)`,
            item: item
        }));
    }

    // Filter out items already selected in currentEmrAssets; mark ones reserved on a letter
    const alreadySelected = new Set(currentEmrAssets.map(a => safeLower(a.serial_no)));
    availableItems = availableItems.filter(entry => !alreadySelected.has(safeLower(entry.serial_no)));
    const reservedMap = typeof getReservedSerials === 'function' ? getReservedSerials() : new Map();
    availableItems.forEach(entry => {
        const dl = reservedMap.get(safeLower(entry.serial_no));
        if (dl) entry.label += `  [reserved: ${dl}]`;
    });

    availableItems.sort((a, b) => safeStr(a.serial_no).localeCompare(safeStr(b.serial_no)));

    availableItems.forEach(entry => {
        const opt = document.createElement('option');
        opt.value = entry.serial_no;
        opt.textContent = entry.label;
        serialSelect.appendChild(opt);
    });

    if (preselectSerial && !alreadySelected.has(safeLower(preselectSerial))) {
        serialSelect.value = preselectSerial;
    }

    updateGatePassPreviewCard();
}

function updateGatePassPreviewCard() {
    const catElem = document.getElementById('gp-category');
    const serElem = document.getElementById('gp-serial-select');
    const previewCard = document.getElementById('gp-item-preview-card');
    if (!previewCard || !catElem || !serElem) return;

    const category = catElem.value;
    const serialNo = serElem.value;

    if (!serialNo) {
        previewCard.style.display = 'none';
        return;
    }

    let found = null;
    if (category === 'DU') {
        found = (db.inventory.pressure_dus || []).find(i => safeLower(i.serial_no) === safeLower(serialNo));
    } else {
        found = (db.inventory.stps || []).find(i => safeLower(i.serial_no) === safeLower(serialNo));
    }

    if (found) {
        previewCard.style.display = 'block';
        document.getElementById('gp-preview-title').textContent = `${found.make || 'HPCL'} - ${found.model || 'Standard Model'}`;
        document.getElementById('gp-preview-details').textContent = `${category === 'DU' ? found.variant : found.hp_type + ' HP'} | Receipt: ${formatDisplayDate(found.receipt_date) || 'N/A'}`;
        document.getElementById('gp-preview-serial').textContent = found.serial_no;
    } else {
        previewCard.style.display = 'none';
    }
}

// Core EMR creation, shared by "Issue EMR" and "Record dispatch" on a dispatch letter.
// Moves each asset from stock to the outward ledger and stores one EMR voucher.
// Returns { emr, processed, skipped } or throws an Error with a user-facing message.
function createEmr(opts) {
    const emrNumber = safeStr(opts.emrNumber);
    if (!emrNumber || !opts.outDate || !safeStr(opts.vendor) || !safeStr(opts.destination)) {
        throw new Error('Fill in EMR no, out date, vendor and destination.');
    }
    if ((db.gate_passes || []).some(g => safeUpper(g.emr_no || g.gate_pass_no) === safeUpper(emrNumber))) {
        throw new Error(`EMR no '${emrNumber}' is already used. Enter a different number.`);
    }
    const outDate = formatDisplayDate(opts.outDate);
    const gpReference = safeStr(opts.gatePassNo);
    const processedItems = [];
    const skipped = [];

    (opts.assets || []).forEach(asset => {
        const isDu = asset.category === 'DU';
        const list = isDu ? db.inventory.pressure_dus : db.inventory.stps;
        const idx = list.findIndex(i => safeLower(i.serial_no) === safeLower(asset.serial_no));
        if (idx === -1) { skipped.push(asset.serial_no); return; }
        const itemObj = list[idx];
        list.splice(idx, 1);
        // Snapshot taken now, so a later cancel restores today's condition/remarks
        const snap = buildEmrAsset(asset.category, itemObj);
        const common = {
            serial_no: itemObj.serial_no,
            make: itemObj.make || '',
            model: itemObj.model || '',
            status: itemObj.status || '',
            receipt_date: itemObj.receipt_date || '',
            gate_pass_date: outDate,
            gate_pass_no: gpReference || emrNumber,
            emr_date: outDate,
            emr_no: emrNumber,
            destination: safeStr(opts.destination),
            vendor: safeStr(opts.vendor),
            out_date: outDate
        };
        if (isDu) {
            const nextSno = db.du_outword.reduce((m, d) => Math.max(m, Number(d.sno) || 0), 0) + 1;
            db.du_outword.push({ sno: nextSno, type: itemObj.type || 'Pressure', variant: itemObj.variant || 'MPD', ...common });
        } else {
            const nextSno = db.stp_outword.reduce((m, d) => Math.max(m, Number(d.sno) || 0), 0) + 1;
            db.stp_outword.push({ sno: nextSno, hp_type: snap.hp_type || '0.75', ...common });
        }
        processedItems.push(snap);
    });

    (opts.otherItems || []).forEach(line => {
        if (!safeStr(line.material)) return;
        const nextSno = db.others_outword.reduce((m, d) => Math.max(m, Number(d.sno) || 0), 0) + 1;
        db.others_outword.push({
            sno: nextSno, material: safeStr(line.material), qty: Number(line.qty) || 1,
            gate_pass_date: outDate, gate_pass_no: gpReference || emrNumber,
            emr_date: outDate, emr_no: emrNumber,
            destination: safeStr(opts.destination), vendor: safeStr(opts.vendor), out_date: outDate
        });
    });

    if (processedItems.length === 0 && !(opts.otherItems || []).some(l => safeStr(l.material))) {
        throw new Error('None of the selected assets are in store any more, so nothing was dispatched.');
    }

    const categoriesSet = new Set(processedItems.map(p => p.category));
    const emrRecord = {
        emr_no: emrNumber,
        gate_pass_no: gpReference || emrNumber,
        category: categoriesSet.size === 1 ? processedItems[0].category : (categoriesSet.size ? 'DU + STP' : 'OTHERS'),
        serial_no: processedItems.map(p => p.serial_no).join(', '),
        make: Array.from(new Set(processedItems.map(p => p.make).filter(Boolean))).join(', ') || 'HPCL',
        model: Array.from(new Set(processedItems.map(p => p.model).filter(Boolean))).join(', '),
        variant: processedItems.map(p => p.variant).join(', '),
        vendor: safeStr(opts.vendor),
        destination: safeStr(opts.destination),
        out_date: outDate,
        vehicle_no: safeStr(opts.vehicleNo),
        driver: safeStr(opts.driver),
        remarks: safeStr(opts.remarks),
        dl_no: safeStr(opts.dlNo),
        dispatched_by: safeStr(opts.dispatchedBy),
        created_at: new Date().toISOString(),
        items: processedItems,
        other_items: (opts.otherItems || []).filter(l => safeStr(l.material)),
        photos: []
    };
    db.gate_passes.push(emrRecord);
    return { emr: emrRecord, processed: processedItems, skipped };
}

async function handleGatePassSubmit() {
    // A serial picked in the dropdown but not yet added is included automatically
    const pendingSerial = document.getElementById('gp-serial-select') ? safeStr(document.getElementById('gp-serial-select').value) : '';
    const pendingCategory = document.getElementById('gp-category') ? document.getElementById('gp-category').value : 'DU';
    if (pendingSerial && !currentEmrAssets.some(a => safeLower(a.serial_no) === safeLower(pendingSerial))) {
        const list = pendingCategory === 'DU' ? db.inventory.pressure_dus : db.inventory.stps;
        const found = list.find(i => safeLower(i.serial_no) === safeLower(pendingSerial));
        if (found) {
            currentEmrAssets.push(buildEmrAsset(pendingCategory, found));
            renderEmrSelectedAssetsTable();
        }
    }

    if (currentEmrAssets.length === 0) {
        alert('Add at least one asset to this EMR.');
        return;
    }

    const reserved = typeof getReservedSerials === 'function' ? getReservedSerials() : new Map();
    const onLetters = currentEmrAssets.filter(a => reserved.has(safeLower(a.serial_no)));
    if (onLetters.length && !confirm(`${onLetters.map(a => `${a.serial_no} (${reserved.get(safeLower(a.serial_no))})`).join(', ')} ${onLetters.length === 1 ? 'is' : 'are'} reserved on a dispatch letter.\n\nIssue this EMR anyway? Use "Record dispatch" on the letter instead to keep the paper trail linked.`)) {
        return;
    }

    let result;
    try {
        result = createEmr({
            assets: currentEmrAssets,
            emrNumber: document.getElementById('gp-number').value,
            outDate: document.getElementById('gp-out-date').value,
            vendor: document.getElementById('gp-vendor').value,
            destination: document.getElementById('gp-destination').value,
            gatePassNo: document.getElementById('gp-emr-no').value,
            vehicleNo: document.getElementById('gp-vehicle-no').value,
            driver: document.getElementById('gp-driver').value,
            remarks: document.getElementById('gp-remarks').value
        });
    } catch (err) {
        alert(err.message);
        return;
    }

    if (typeof commitCapturedPhotos === 'function') {
        await commitCapturedPhotos('gp', result.emr);
    }

    currentEmrAssets = [];
    saveState();
    refreshAllViews();

    const modal = document.getElementById('gate-pass-modal');
    if (modal) modal.classList.remove('active');

    showToast(`EMR ${result.emr.emr_no} issued for ${result.processed.length} asset(s).`);
    if (typeof uploadPendingPhotos === 'function') uploadPendingPhotos();
    openGatePassPrintModal(result.emr.emr_no);
}

function findGatePassIndex(voucherNumber) {
    const key = safeLower(voucherNumber);
    let idx = (db.gate_passes || []).findIndex(g => safeLower(g.emr_no) === key);
    if (idx === -1) idx = (db.gate_passes || []).findIndex(g => safeLower(g.gate_pass_no) === key);
    return idx;
}

function revertGatePass(voucherNumber) {
    const gpIndex = findGatePassIndex(voucherNumber);
    if (gpIndex === -1) {
        alert(`EMR '${voucherNumber}' not found.`);
        return;
    }

    const gp = db.gate_passes[gpIndex];
    const displayNum = gp.emr_no || gp.gate_pass_no;

    const itemsToRevert = (gp.items && gp.items.length > 0) ? gp.items : [{
        category: gp.category,
        serial_no: gp.serial_no,
        make: gp.make,
        model: gp.model,
        variant: gp.variant
    }];

    const serialListStr = itemsToRevert.map(i => i.serial_no).join(', ');

    const photoNote = Array.isArray(gp.photos) && gp.photos.length ? `\nIts ${gp.photos.length} photo(s) will be deleted from this device.` : '';
    const letterNote = gp.dl_no ? `\nDispatch letter ${gp.dl_no} goes back to "waiting at depot".` : '';
    if (!confirm(`Cancel EMR '${displayNum}'?\n\n${itemsToRevert.length} asset(s) (${serialListStr}) go back to in-store stock and are removed from the outward ledger.${letterNote}${photoNote}`)) {
        return;
    }

    db.gate_passes.splice(gpIndex, 1);

    // A letter dispatched through this EMR goes back to "waiting at depot"
    if (gp.dl_no && Array.isArray(db.dispatch_letters)) {
        const letter = db.dispatch_letters.find(l => safeLower(l.dl_no) === safeLower(gp.dl_no));
        if (letter && letter.status === 'dispatched') {
            letter.status = 'issued';
            ['emr_no', 'dispatched_on', 'dispatched_by', 'vehicle_no', 'driver', 'gate_pass_no', 'dispatch_remarks'].forEach(k => delete letter[k]);
        }
    }
    // Its photos are removed from this device
    if (Array.isArray(gp.photos) && gp.photos.length && typeof deletePhotosLocal === 'function') {
        deletePhotosLocal(gp.photos.map(p => p.id));
    }
    // Other-material lines issued on this EMR
    db.others_outword = db.others_outword.filter(o => !(safeLower(o.emr_no) === safeLower(displayNum) && Array.isArray(gp.other_items) && gp.other_items.length));

    const emrKey = safeLower(displayNum);
    const gpKey = safeLower(gp.gate_pass_no);
    // Pick the outward row for this serial that belongs to THIS EMR (a serial can be dispatched more than once over time)
    const takeOutwardRow = (list, serial) => {
        let idx = list.findIndex(r => safeLower(r.serial_no) === safeLower(serial) &&
            (safeLower(r.emr_no) === emrKey || (gpKey && safeLower(r.gate_pass_no) === gpKey)));
        if (idx === -1) {
            // Legacy rows without EMR numbers: fall back to the most recent row for that serial
            for (let i = list.length - 1; i >= 0; i--) {
                if (safeLower(list[i].serial_no) === safeLower(serial) && !safeStr(list[i].emr_no)) { idx = i; break; }
            }
        }
        return idx !== -1 ? list.splice(idx, 1)[0] : null;
    };
    const returnNote = `Returned to store from cancelled EMR ${displayNum}`;
    const joinRemarks = (orig) => orig ? `${orig} | ${returnNote}` : returnNote;

    let restored = 0;
    itemsToRevert.forEach(item => {
        if (!safeStr(item.serial_no)) return;
        const isDu = item.category === 'DU' || (item.category && item.category.includes('DU'));
        if (isDu) {
            const outRow = takeOutwardRow(db.du_outword, item.serial_no) || {};
            if (db.inventory.pressure_dus.some(d => safeLower(d.serial_no) === safeLower(item.serial_no))) return;
            const nextSno = (db.inventory.pressure_dus || []).reduce((max, d) => Number(d.sno) > max ? Number(d.sno) : max, 0) + 1;
            db.inventory.pressure_dus.push({
                sno: nextSno,
                type: item.type || outRow.type || 'Pressure',
                variant: item.variant || outRow.variant || 'MPD',
                make: item.make || outRow.make || 'UNKNOWN',
                model: item.model || outRow.model || '',
                serial_no: item.serial_no,
                receipt_date: item.receipt_date || outRow.receipt_date || todayLocal(),
                status: item.status || outRow.status || 'N',
                remarks: joinRemarks(item.remarks)
            });
            restored++;
        } else {
            const outRow = takeOutwardRow(db.stp_outword, item.serial_no) || {};
            if (db.inventory.stps.some(s => safeLower(s.serial_no) === safeLower(item.serial_no))) return;
            const nextSno = (db.inventory.stps || []).reduce((max, s) => Number(s.sno) > max ? Number(s.sno) : max, 0) + 1;
            db.inventory.stps.push({
                sno: nextSno,
                hp_type: safeStr(item.hp_type || item.variant || outRow.hp_type || '0.75').replace(/\s*HP$/i, ''),
                make: item.make || outRow.make || 'UNKNOWN',
                serial_no: item.serial_no,
                model: item.model || outRow.model || '',
                receipt_date: item.receipt_date || outRow.receipt_date || todayLocal(),
                status: item.status || outRow.status || 'N',
                remarks: joinRemarks(item.remarks)
            });
            restored++;
        }
    });

    saveState();
    refreshAllViews();
    populateFilterDropdowns();

    showToast(`EMR ${displayNum} cancelled. ${restored} asset(s) returned to Store Inventory.`);
}

function openGatePassPrintModal(voucherNumber) {
    const gpIdx = findGatePassIndex(voucherNumber);
    if (gpIdx === -1) return;
    let gp = db.gate_passes[gpIdx];

    const modal = document.getElementById('gate-pass-print-modal');
    const container = document.getElementById('printable-voucher-body');
    if (!container || !modal) return;
    const printTitle = document.getElementById('print-modal-title');
    if (printTitle) printTitle.textContent = 'EMR voucher';

    const rawEmr = safeStr(gp.emr_no || gp.gate_pass_no);
    const qrMockCode = `HPCL-DESUR-EMR-${rawEmr.replace(/[^a-zA-Z0-9]/g, '')}`;
    const rawItems = (gp.items && gp.items.length > 0) ? gp.items : null;
    gp = escapeRecord(gp);
    const emrDisplay = gp.emr_no || gp.gate_pass_no;
    const gpDisplay = gp.gate_pass_no || gp.gp_no || 'N/A';

    const itemsToPrint = rawItems ? rawItems.map(escapeRecord) : [{
        category: gp.category,
        serial_no: gp.serial_no,
        make: gp.make,
        model: gp.model,
        variant: gp.variant
    }];

    const otherLines = Array.isArray(db.gate_passes[gpIdx].other_items) ? db.gate_passes[gpIdx].other_items.map(escapeRecord) : [];
    const equipmentRowsHtml = itemsToPrint.filter(it => it.serial_no).map((it, idx) => {
        const isDu = it.category === 'DU' || (it.category && it.category.includes('DU'));
        const desc = isDu ? 'Pressure Dispensing Unit (DU)' : 'Submersible Turbine Pump (STP)';
        const spec = it.variant || (it.hp_type ? `${it.hp_type} HP` : 'Standard');
        return `
            <tr>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align:center; font-weight:700;">${idx + 1}</td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1;">
                    <strong style="font-size: 0.95rem;">${desc}</strong>
                    <div style="font-size: 0.8rem; color: #475569; margin-top: 2px;">Specification: ${spec}</div>
                </td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1;">
                    <strong>${it.make || 'HPCL Depot'}</strong><br>
                    <span style="font-size: 0.8rem; color: #475569;">${it.model || 'Standard Model'}</span>
                </td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1; font-family: 'Barlow Condensed', sans-serif; font-weight: 800; font-size: 1.05rem; color: #D2232A;">
                    ${it.serial_no}
                </td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align:center; font-weight:800; font-size: 0.95rem;">
                    1 Unit
                </td>
            </tr>
        `;
    }).join('') + otherLines.map((o, i) => `
            <tr>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align:center; font-weight:700;">${itemsToPrint.filter(it => it.serial_no).length + i + 1}</td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1;" colspan="3"><strong>${o.material}</strong></td>
                <td style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align:center; font-weight:800;">${o.qty || 1}</td>
            </tr>`).join('');

    const footerTotalHtml = `
        <tr style="background: #f8fafc; font-weight: 700;">
            <td colspan="4" style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align: right; color: #334155; font-size: 0.9rem;">
                Total Equipment Units on this EMR:
            </td>
            <td style="padding: 10px 12px; border: 1px solid #cbd5e1; text-align: center; color: #004899; font-size: 1.05rem; font-weight: 800;">
                ${itemsToPrint.filter(it => it.serial_no).length} Unit${itemsToPrint.filter(it => it.serial_no).length === 1 ? '' : 's'}${otherLines.length ? ` + ${otherLines.length} other line${otherLines.length === 1 ? '' : 's'}` : ''}
            </td>
        </tr>
    `;

    container.innerHTML = `
        <div class="printable-pass-sheet" style="font-family: 'Barlow', Arial, sans-serif; color: #0f172a; line-height: 1.45;">
            <!-- Header with Official Crest & Dual Language -->
            <div style="border-bottom: 2.5px solid #004899; padding-bottom: 1.25rem; margin-bottom: 1.25rem; display: flex; justify-content: space-between; align-items: center;">
                <div style="display:flex; align-items:center; gap: 14px;">
                    <img src="HPCL.png" alt="HPCL Crest" style="height: 52px; width: auto; object-fit: contain;">
                    <div>
                        <h3 style="margin: 0; font-size: 1.05rem; color: #D2232A; font-weight: 700; font-family: 'Noto Sans Devanagari', sans-serif;">हिन्दुस्तान पेट्रोलियम कॉर्पोरेशन लिमिटेड</h3>
                        <h2 style="margin: 2px 0 0 0; font-size: 1.35rem; color: #004899; text-transform: uppercase; font-weight: 800; font-family: 'Barlow Condensed', 'Arial Narrow', sans-serif;">HINDUSTAN PETROLEUM CORPORATION LIMITED</h2>
                        <p style="margin: 3px 0 0 0; font-size: 0.85rem; color: #475569; font-weight: 600;">Desur IRD Depot, Belgaum Regional Office (RRO) - 590014</p>
                    </div>
                </div>
                <div style="text-align: right;">
                    <span style="display: inline-block; background: #004899; color: #ffffff; padding: 4px 14px; border-radius: 4px; font-weight: 800; font-size: 0.88rem; letter-spacing: 0.05em;">EMR CHALLAN</span>
                    <p style="margin: 5px 0 0 0; font-size: 1rem; font-family: 'Barlow Condensed', sans-serif; font-weight: 800; color: #D2232A;">${emrDisplay}</p>
                    <span style="font-size: 0.72rem; color: #64748b; font-family: 'Barlow Condensed', sans-serif;">CODE: ${qrMockCode}</span>
                </div>
            </div>

            <!-- Dispatch Transport Details Grid -->
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.25rem; margin-bottom: 1.25rem; font-size: 0.85rem; background: #f8fafc; padding: 1.15rem; border-radius: 6px; border: 1px solid #cbd5e1;">
                <div>
                    <p style="margin: 0 0 6px 0;"><strong>Out Date:</strong> <span style="font-weight:700;">${formatDisplayDate(gp.out_date) || '-'}</span></p>
                    ${gp.dl_no ? `<p style="margin: 0 0 6px 0;"><strong>Against Dispatch Letter:</strong> <span style="font-weight:700;">${gp.dl_no}</span></p>` : ''}
                    <p style="margin: 0 0 6px 0;"><strong>Authorized Transporter / Vendor:</strong> <br><span style="color: #004899; font-weight:800; font-size: 0.95rem;">${gp.vendor || '-'}</span></p>
                    <p style="margin: 0;"><strong>Destination (RO Outlet / Site):</strong> <br><span style="font-weight:700;">${gp.destination || '-'}</span></p>
                </div>
                <div>
                    <p style="margin: 0 0 6px 0;"><strong>Gate Pass / Security Slip Ref:</strong> <span style="font-family: 'Barlow Condensed', sans-serif; font-weight:700;">${gpDisplay}</span></p>
                    <p style="margin: 0 0 6px 0;"><strong>Vehicle Reg. Number:</strong> <span style="font-family: 'Barlow Condensed', sans-serif; font-weight:700; color:#ea580c;">${gp.vehicle_no || 'N/A'}</span></p>
                    <p style="margin: 0 0 6px 0;"><strong>Driver Name &amp; Contact:</strong> <span style="font-weight:700;">${gp.driver || 'N/A'}</span></p>
                    <p style="margin: 0;"><strong>Dispatch Photos:</strong> <span style="font-weight:700;">${Array.isArray(db.gate_passes[gpIdx].photos) && db.gate_passes[gpIdx].photos.length ? db.gate_passes[gpIdx].photos.length + ' on record' : 'Not yet taken'}</span></p>
                </div>
            </div>

            <!-- Stores & Equipment Table -->
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 1.25rem; font-size: 0.85rem; border: 1.5px solid #0f172a;">
                <thead>
                    <tr style="background: #0f172a; color: #ffffff; text-align: left;">
                        <th style="padding: 9px 12px; border: 1px solid #334155; width: 45px; text-align:center;">#</th>
                        <th style="padding: 9px 12px; border: 1px solid #334155;">Description of Stores</th>
                        <th style="padding: 9px 12px; border: 1px solid #334155;">Make &amp; Model</th>
                        <th style="padding: 9px 12px; border: 1px solid #334155;">Serial Number</th>
                        <th style="padding: 9px 12px; border: 1px solid #334155; text-align:center; width: 90px;">Quantity</th>
                    </tr>
                </thead>
                <tbody>
                    ${equipmentRowsHtml}
                    ${footerTotalHtml}
                </tbody>
            </table>

            <!-- Remarks & Notes -->
            ${gp.remarks ? `<div style="font-size: 0.82rem; background: #fff7ed; border-left: 4px solid #D2232A; padding: 8px 12px; margin-bottom: 1.25rem;"><strong>Remarks:</strong> ${gp.remarks}</div>` : ''}

            <div style="font-size: 0.74rem; color: #64748b; background: #f1f5f9; padding: 8px 12px; border-radius: 4px; margin-bottom: 2.5rem; line-height: 1.4;">
                <strong>Statutory Notice:</strong> The materials specified above belong to HPCL Desur IRD Depot and are dispatched under authorized EMR custody. Transporter and receiving dealer must verify serial numbers upon arrival and adhere strictly to OISD / PESO safety compliance standards.
            </div>

            <!-- 4 Formal Signature Blocks -->
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 1rem; text-align: center; font-size: 0.78rem;">
                <div style="border-top: 1.5px dashed #475569; padding-top: 8px;">
                    <strong>Issued / Dispatched By</strong><br>
                    <span style="color: #475569;">Vishal<br>(Junior Executive - RE)</span>
                </div>
                <div style="border-top: 1.5px dashed #475569; padding-top: 8px;">
                    <strong>Checked &amp; Verified By</strong><br>
                    <span style="color: #475569;">Vinita Kumari<br>(Manager - RE)</span>
                </div>
                <div style="border-top: 1.5px dashed #475569; padding-top: 8px;">
                    <strong>Approved By</strong><br>
                    <span style="color: #475569;">Praise K James<br>(Manager - RE Head)</span>
                </div>
                <div style="border-top: 1.5px dashed #475569; padding-top: 8px;">
                    <strong>Security Gate Out</strong><br>
                    <span style="color: #475569;">Stamp &amp; Time Out<br>(Main Gate)</span>
                </div>
            </div>
        </div>
    `;

    modal.classList.add('active');
}

// Window global exports & aliases
window.quickIssueEMR = quickIssueGatePass;
window.quickIssueGatePass = quickIssueGatePass;
window.openEMRModal = openGatePassModal;
window.openGatePassModal = openGatePassModal;
window.openEMRPrintModal = openGatePassPrintModal;
window.openGatePassPrintModal = openGatePassPrintModal;
window.revertEMR = revertGatePass;
window.revertGatePass = revertGatePass;
window.renderEMRView = renderGatePassView;
window.renderGatePassView = renderGatePassView;
window.addAssetToEmr = addAssetToEmr;
window.removeAssetFromEmr = removeAssetFromEmr;
window.copyToClipboard = copyToClipboard;

// ==========================================================================
// PWA Service Worker & Installation
// ==========================================================================
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then((reg) => console.log('PWA Service Worker registered:', reg.scope))
            .catch((err) => console.warn('PWA Service Worker registration failed:', err));
    });
}

// ==========================================================================
// Automatic PWA Install Popup on App Launch (Any Device)
// ==========================================================================
let deferredPrompt = null;
let pwaPromptShown = false;

// 1. Capture beforeinstallprompt if available (Chrome, Edge, Android)
window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    console.log('beforeinstallprompt event captured');
    triggerInstallPopup();
});

// 2. Check if already running standalone
function isRunningStandalone() {
    return (
        window.matchMedia('(display-mode: standalone)').matches ||
        window.navigator.standalone === true ||
        document.referrer.includes('android-app://')
    );
}

// 3. Detect device guidance
function getDeviceInstallGuidance() {
    const ua = navigator.userAgent || navigator.vendor || window.opera || '';
    const isIOS = /iPad|iPhone|iPod/.test(ua) && !window.MSStream;
    const isAndroid = /android/i.test(ua);

    if (isIOS) {
        return {
            type: 'ios',
            buttonHtml: '<i class="fa-brands fa-apple"></i> <span>Add to Home Screen</span>',
            hintText: 'On iOS: Tap the <strong>Share</strong> button <i class="fa-solid fa-arrow-up-from-bracket"></i> in Safari, then scroll down and tap <strong>"Add to Home Screen"</strong> <i class="fa-solid fa-square-plus"></i>.'
        };
    } else if (isAndroid) {
        return {
            type: 'android',
            buttonHtml: '<i class="fa-brands fa-android"></i> <span>Install App Now</span>',
            hintText: 'Tap <strong>Install App Now</strong> to add HPCL Portal directly to your Android home screen.'
        };
    } else {
        return {
            type: 'desktop',
            buttonHtml: '<i class="fa-solid fa-download"></i> <span>Install Desktop App</span>',
            hintText: 'Click <strong>Install Desktop App</strong> to add HPCL Portal to your desktop for quick launch and offline access.'
        };
    }
}

// 4. Trigger the popup automatically when opening the application
function triggerInstallPopup() {
    if (pwaPromptShown) return;
    if (isRunningStandalone()) {
        console.log('App is running in standalone mode; skipping install popup.');
        return;
    }

    if (sessionStorage.getItem('hpcl_pwa_dismissed') === 'true') {
        return;
    }

    const modal = document.getElementById('pwa-install-modal');
    if (!modal) return;

    pwaPromptShown = true;

    // Tailor hint and button according to device
    const guidance = getDeviceInstallGuidance();
    const hintTextEl = document.getElementById('pwa-hint-text');
    const confirmBtn = document.getElementById('pwa-confirm-install-btn');
    if (hintTextEl) hintTextEl.innerHTML = guidance.hintText;
    if (confirmBtn) confirmBtn.innerHTML = guidance.buttonHtml;

    setTimeout(() => {
        modal.classList.add('active');
    }, 700);
}

// Setup click handlers for the install modal
function setupPwaInstallModalHandlers() {
    const modal = document.getElementById('pwa-install-modal');
    const closeBtn = document.getElementById('pwa-modal-close');
    const dismissBtn = document.getElementById('pwa-dismiss-btn');
    const confirmBtn = document.getElementById('pwa-confirm-install-btn');

    const closeModal = () => {
        if (modal) modal.classList.remove('active');
        sessionStorage.setItem('hpcl_pwa_dismissed', 'true');
    };

    if (closeBtn) closeBtn.addEventListener('click', closeModal);
    if (dismissBtn) dismissBtn.addEventListener('click', closeModal);

    if (confirmBtn) {
        confirmBtn.addEventListener('click', async () => {
            if (deferredPrompt) {
                try {
                    deferredPrompt.prompt();
                    const { outcome } = await deferredPrompt.userChoice;
                    console.log('User choice outcome:', outcome);
                    deferredPrompt = null;
                } catch (err) {
                    console.warn('Error during install prompt:', err);
                }
                closeModal();
            } else {
                const ua = navigator.userAgent || '';
                const isIOS = /iPad|iPhone|iPod/.test(ua) && !window.MSStream;
                if (isIOS) {
                    alert("📲 How to Install on iPhone / iPad:\n\n1. Tap the 'Share' icon (square with arrow up) at the bottom of Safari.\n2. Scroll down and tap 'Add to Home Screen'.\n3. Tap 'Add' in the top right corner.");
                } else {
                    alert("📲 How to Install HPCL App:\n\n• In Chrome/Edge: Look for the 'Install' icon in your browser address bar (top right), or click Menu (⋮) -> 'Install HPCL Active Inventory'.\n• On Mobile: Tap Menu (⋮) -> 'Install app' or 'Add to Home screen'.");
                }
                closeModal();
            }
        });
    }

    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });
    }
}

// Initialize handler and trigger popup on app start
setupPwaInstallModalHandlers();

window.addEventListener('load', () => {
    setTimeout(() => {
        triggerInstallPopup();
    }, 900);
});


