/* ==========================================================================
   Dispatch letters, depot dispatch with photo proof, photo storage & upload.
   Depends on globals from app.js (db, saveState, refreshAllViews, createEmr,
   buildEmrAsset, escapeHtml, safeStr, safeLower, todayLocal, showToast, ...).
   ========================================================================== */

// --------------------------------------------------------------------------
// Photo storage (IndexedDB). Photos are too big for localStorage, so only
// their metadata lives in db (emr.photos[]); the JPEG blobs live here.
// --------------------------------------------------------------------------
const PHOTO_DB_NAME = 'hpcl_stores_photos';
const PHOTO_STORE = 'photos';
let photoDbPromise = null;

function openPhotoDb() {
    if (photoDbPromise) return photoDbPromise;
    photoDbPromise = new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) { reject(new Error('This browser cannot store photos (no IndexedDB).')); return; }
        const req = indexedDB.open(PHOTO_DB_NAME, 1);
        req.onupgradeneeded = () => {
            const idb = req.result;
            if (!idb.objectStoreNames.contains(PHOTO_STORE)) idb.createObjectStore(PHOTO_STORE, { keyPath: 'id' });
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
    return photoDbPromise;
}

function photoTx(mode, fn) {
    return openPhotoDb().then(idb => new Promise((resolve, reject) => {
        const tx = idb.transaction(PHOTO_STORE, mode);
        const store = tx.objectStore(PHOTO_STORE);
        let result;
        Promise.resolve(fn(store, r => { result = r; })).catch(reject);
        tx.oncomplete = () => resolve(result);
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error);
    }));
}

function photoPut(rec) { return photoTx('readwrite', store => { store.put(rec); }); }
function photoGet(id) {
    return photoTx('readonly', (store, done) => {
        const r = store.get(id);
        r.onsuccess = () => done(r.result || null);
    });
}
function photoDelete(id) { return photoTx('readwrite', store => { store.delete(id); }); }
function photoAllIds() {
    return photoTx('readonly', (store, done) => {
        const r = store.getAllKeys();
        r.onsuccess = () => done(r.result || []);
    });
}

const photoUrlCache = new Map();

// Best available source for a photo: the local blob, else the Drive copy.
async function photoSrc(meta) {
    if (!meta) return null;
    if (photoUrlCache.has(meta.id)) return photoUrlCache.get(meta.id);
    try {
        const rec = await photoGet(meta.id);
        if (rec && rec.blob) {
            const url = URL.createObjectURL(rec.blob);
            photoUrlCache.set(meta.id, url);
            return url;
        }
    } catch (e) { /* fall through to Drive */ }
    if (meta.drive_id) return `https://drive.google.com/thumbnail?id=${encodeURIComponent(meta.drive_id)}&sz=w1600`;
    return null;
}

function deletePhotosLocal(ids) {
    (ids || []).forEach(id => {
        if (photoUrlCache.has(id)) { URL.revokeObjectURL(photoUrlCache.get(id)); photoUrlCache.delete(id); }
        photoDelete(id).catch(() => {});
    });
}

function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
        const fr = new FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = () => reject(fr.error);
        fr.readAsDataURL(blob);
    });
}

async function dataUrlToBlob(dataUrl) {
    const res = await fetch(dataUrl);
    return res.blob();
}

async function exportPhotosForBackup() {
    const out = {};
    const ids = await photoAllIds().catch(() => []);
    for (const id of ids) {
        const rec = await photoGet(id);
        if (rec && rec.blob) out[id] = { data: await blobToDataUrl(rec.blob), taken_at: rec.taken_at };
    }
    return out;
}

async function importPhotosFromBackup(map) {
    let n = 0;
    for (const id of Object.keys(map || {})) {
        try {
            const blob = await dataUrlToBlob(map[id].data);
            await photoPut({ id, blob, taken_at: map[id].taken_at || '' });
            n++;
        } catch (e) { console.warn('Photo restore failed', id, e); }
    }
    return n;
}

// --------------------------------------------------------------------------
// Photo processing: shrink to 1600px, stamp depot / date / time on the image.
// --------------------------------------------------------------------------
function loadImageFromFile(file) {
    return new Promise((resolve, reject) => {
        const url = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
        img.onerror = () => { URL.revokeObjectURL(url); reject(new Error(`${file.name || 'This file'} is not an image this browser can open.`)); };
        img.src = url;
    });
}

function formatStampTime(d) {
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

async function processPhoto(file) {
    const img = await loadImageFromFile(file);
    const MAX = 1600;
    const scale = Math.min(1, MAX / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.round(img.naturalWidth * scale);
    const h = Math.round(img.naturalHeight * scale);
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, w, h);

    const now = new Date();
    const stamp = `HPCL Desur IRD  ·  ${formatStampTime(now)}`;
    const fontPx = Math.max(16, Math.round(w / 42));
    ctx.font = `600 ${fontPx}px Barlow, Arial, sans-serif`;
    const pad = Math.round(fontPx * 0.55);
    const bandH = fontPx + pad * 2;
    ctx.fillStyle = 'rgba(15, 42, 74, 0.78)';
    ctx.fillRect(0, h - bandH, w, bandH);
    ctx.fillStyle = '#F5B700';
    ctx.fillRect(0, h - bandH, Math.max(6, Math.round(fontPx / 3)), bandH);
    ctx.fillStyle = '#FFFFFF';
    ctx.textBaseline = 'middle';
    ctx.fillText(stamp, pad * 2, h - bandH / 2);

    const blob = await new Promise(res => canvas.toBlob(res, 'image/jpeg', 0.78));
    if (!blob) throw new Error('Could not process the photo.');
    return { blob, width: w, height: h, taken_at: now.toISOString() };
}

function newPhotoId() {
    return 'ph_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
}

// --------------------------------------------------------------------------
// Photo capture widget. Any element with data-photo-capture="<key>" becomes
// "Take photo" + "Choose from gallery" with a preview grid.
// --------------------------------------------------------------------------
const photoCaptures = {};

function initPhotoCaptures() {
    document.querySelectorAll('[data-photo-capture]').forEach(root => {
        const key = root.getAttribute('data-photo-capture');
        const required = root.getAttribute('data-required') === '1';
        root.innerHTML = `
            <div class="photo-capture-actions">
                <label class="btn btn-outline">
                    <i class="fa-solid fa-camera"></i> Take photo
                    <input type="file" accept="image/*" capture="environment" hidden data-cap-input>
                </label>
                <label class="btn btn-ghost">
                    <i class="fa-regular fa-images"></i> Choose from gallery
                    <input type="file" accept="image/*" multiple hidden data-cap-input>
                </label>
                <span class="text-muted" data-cap-count style="font-size:.88rem;"></span>
            </div>
            <div class="hint">${required ? 'At least one photo is needed. ' : ''}Show the vehicle number plate and the serial plate of each unit. Date and time are stamped on the photo.</div>
            <div class="photo-grid" data-cap-grid></div>`;
        photoCaptures[key] = { root, required, pending: [] };

        root.querySelectorAll('[data-cap-input]').forEach(input => {
            input.addEventListener('change', async () => {
                const files = Array.from(input.files || []);
                input.value = '';
                for (const f of files) {
                    try {
                        const processed = await processPhoto(f);
                        processed.tmpId = newPhotoId();
                        processed.url = URL.createObjectURL(processed.blob);
                        photoCaptures[key].pending.push(processed);
                    } catch (err) {
                        alert(err.message);
                    }
                }
                renderCapture(key);
            });
        });

        root.addEventListener('click', (e) => {
            const rm = e.target.closest('[data-cap-remove]');
            if (rm) {
                const id = rm.getAttribute('data-cap-remove');
                const cap = photoCaptures[key];
                const p = cap.pending.find(x => x.tmpId === id);
                if (p) URL.revokeObjectURL(p.url);
                cap.pending = cap.pending.filter(x => x.tmpId !== id);
                renderCapture(key);
                return;
            }
            const img = e.target.closest('img[data-cap-img]');
            if (img) openLightbox(img.src, 'Not saved yet');
        });
    });
}

function renderCapture(key) {
    const cap = photoCaptures[key];
    if (!cap) return;
    const grid = cap.root.querySelector('[data-cap-grid]');
    grid.innerHTML = cap.pending.map(p => `
        <div class="photo-tile">
            <img src="${p.url}" alt="New dispatch photo" data-cap-img>
            <button type="button" class="remove" data-cap-remove="${p.tmpId}" aria-label="Remove photo"><i class="fa-solid fa-xmark"></i></button>
        </div>`).join('');
    const count = cap.root.querySelector('[data-cap-count]');
    if (count) count.textContent = cap.pending.length ? `${cap.pending.length} photo${cap.pending.length === 1 ? '' : 's'} ready` : '';
    cap.root.classList.remove('required-missing');
}

function resetPhotoCapture(key) {
    const cap = photoCaptures[key];
    if (!cap) return;
    cap.pending.forEach(p => URL.revokeObjectURL(p.url));
    cap.pending = [];
    renderCapture(key);
}

function captureCount(key) {
    return photoCaptures[key] ? photoCaptures[key].pending.length : 0;
}

// Saves the captured photos for an EMR: blobs to IndexedDB, metadata onto emr.photos.
async function commitCapturedPhotos(key, emr) {
    const cap = photoCaptures[key];
    if (!cap || !cap.pending.length || !emr) return [];
    if (!Array.isArray(emr.photos)) emr.photos = [];
    const metas = [];
    for (const p of cap.pending) {
        const id = newPhotoId();
        try {
            await photoPut({ id, blob: p.blob, taken_at: p.taken_at });
            const meta = { id, taken_at: p.taken_at, width: p.width, height: p.height, size: p.blob.size, emr_no: emr.emr_no };
            emr.photos.push(meta);
            metas.push(meta);
            photoUrlCache.set(id, p.url); // reuse the preview URL
        } catch (err) {
            console.error('Saving photo failed', err);
            alert('A photo could not be saved on this device: ' + err.message);
        }
    }
    cap.pending = [];
    renderCapture(key);
    return metas;
}

// --------------------------------------------------------------------------
// Upload photos to Google Drive through the Apps Script Web App
// --------------------------------------------------------------------------
let photoUploadRunning = false;

function pendingPhotoMetas() {
    const out = [];
    (db.gate_passes || []).forEach(g => (g.photos || []).forEach(p => { if (!p.drive_id) out.push({ g, p }); }));
    return out;
}

function updatePendingPhotoCount() {
    const el = document.getElementById('pending-photo-count');
    if (el) el.textContent = pendingPhotoMetas().length;
}

async function uploadPendingPhotos(opts = {}) {
    const url = localStorage.getItem('hpcl_sheets_url') || '';
    const pending = pendingPhotoMetas();
    if (!pending.length) {
        if (opts.manual) showToast('All photos are already uploaded.');
        return;
    }
    if (typeof isAppsScriptUrl !== 'function' || !isAppsScriptUrl(url)) {
        if (opts.manual) alert('Photos stay on this device until an Apps Script Web App URL is saved in Backup & sync.');
        return;
    }
    if (!navigator.onLine) {
        if (opts.manual) alert('You are offline. Photos will upload when the connection is back.');
        return;
    }
    if (photoUploadRunning) return;
    photoUploadRunning = true;
    let done = 0, failed = 0;
    try {
        for (const { g, p } of pending) {
            const rec = await photoGet(p.id).catch(() => null);
            if (!rec || !rec.blob) continue;
            try {
                const dataUrl = await blobToDataUrl(rec.blob);
                const res = await fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify({
                        action: 'upload_photo',
                        photo_id: p.id,
                        emr_no: g.emr_no,
                        dl_no: g.dl_no || '',
                        taken_at: p.taken_at,
                        mime: 'image/jpeg',
                        data: dataUrl.split(',')[1]
                    })
                });
                const json = await res.json();
                if (json && json.status === 'success' && json.file_id) {
                    p.drive_id = json.file_id;
                    p.drive_url = json.url || `https://drive.google.com/file/d/${json.file_id}/view`;
                    done++;
                } else {
                    failed++;
                    console.warn('Photo upload rejected:', json && json.message);
                }
            } catch (err) {
                failed++;
                console.warn('Photo upload failed:', err);
            }
        }
    } finally {
        photoUploadRunning = false;
    }
    if (done) {
        saveState();
        renderGatePassView();
        renderLetters();
    }
    updatePendingPhotoCount();
    if (opts.manual || done) {
        showToast(failed ? `${done} photo(s) uploaded, ${failed} failed. Update the Apps Script to v3 if this keeps happening.` : `${done} photo(s) uploaded to Google Drive.`);
    }
}

// --------------------------------------------------------------------------
// Lightbox
// --------------------------------------------------------------------------
function openLightbox(src, caption) {
    const box = document.getElementById('lightbox');
    if (!box || !src) return;
    document.getElementById('lightbox-img').src = src;
    document.getElementById('lightbox-caption').textContent = caption || '';
    box.classList.add('active');
}
function closeLightbox() {
    const box = document.getElementById('lightbox');
    if (box) box.classList.remove('active');
}

// --------------------------------------------------------------------------
// Dispatch letters
// --------------------------------------------------------------------------
let letterFilter = 'issued';
let letterAssets = [];

function financialYear(dateStr) {
    const iso = dateStr ? (typeof toIsoDate === 'function' ? toIsoDate(dateStr) : dateStr) : '';
    const d = iso ? new Date(iso + 'T00:00:00') : new Date();
    const y = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
    return `${y}-${String((y + 1) % 100).padStart(2, '0')}`;
}

function nextLetterNo(dateStr) {
    const fy = financialYear(dateStr);
    const prefix = `DL/DESUR/${fy}/`;
    let seq = 1;
    (db.dispatch_letters || []).forEach(l => {
        const n = safeStr(l.dl_no);
        if (n.startsWith(prefix)) seq = Math.max(seq, (parseInt(n.slice(prefix.length), 10) || 0) + 1);
    });
    return prefix + String(seq).padStart(3, '0');
}

// serial (lower-case) -> letter no, for every letter still waiting at the depot
function getReservedSerials() {
    const map = new Map();
    (db.dispatch_letters || []).forEach(l => {
        if (l.status !== 'issued') return;
        (l.items || []).forEach(i => { if (safeStr(i.serial_no)) map.set(safeLower(i.serial_no), l.dl_no); });
    });
    return map;
}

function findLetter(dlNo) {
    return (db.dispatch_letters || []).find(l => safeLower(l.dl_no) === safeLower(dlNo));
}

function findEmr(emrNo) {
    return (db.gate_passes || []).find(g => safeLower(g.emr_no) === safeLower(emrNo)) ||
        (db.gate_passes || []).find(g => safeLower(g.gate_pass_no) === safeLower(emrNo));
}

function daysSince(dateStr) {
    const t = safeDateTimestamp(dateStr);
    if (!t) return null;
    const today = new Date(todayIso() + 'T00:00:00').getTime();
    return Math.max(0, Math.round((today - t) / 86400000));
}

function assetsInStockCount(letter) {
    let inStock = 0;
    (letter.items || []).forEach(i => {
        const list = i.category === 'DU' ? db.inventory.pressure_dus : db.inventory.stps;
        if (list.some(x => safeLower(x.serial_no) === safeLower(i.serial_no))) inStock++;
    });
    return inStock;
}

function renderDispatchLine() {
    const letters = db.dispatch_letters || [];
    const fy = financialYear();
    const thisFy = letters.filter(l => l.status !== 'cancelled' && financialYear(l.dl_date) === fy);
    const waiting = letters.filter(l => l.status === 'issued');
    const done = letters.filter(l => l.status === 'dispatched');

    const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    set('dl-count-total', thisFy.length);
    const units = thisFy.reduce((n, l) => n + (l.items || []).length, 0);
    set('dl-note-total', thisFy.length ? `${units} unit${units === 1 ? '' : 's'} released in FY ${fy}` : `None yet in FY ${fy}`);

    set('dl-count-waiting', waiting.length);
    const stage = document.getElementById('dl-stage-waiting');
    if (stage) stage.classList.toggle('is-zero', waiting.length === 0);
    if (waiting.length) {
        const oldest = waiting.slice().sort((a, b) => safeDateTimestamp(a.dl_date) - safeDateTimestamp(b.dl_date))[0];
        const age = daysSince(oldest.dl_date);
        set('dl-note-waiting', `Oldest: ${oldest.dl_no}${age !== null ? `, ${age === 0 ? 'issued today' : age + ' day' + (age === 1 ? '' : 's') + ' old'}` : ''}`);
    } else {
        set('dl-note-waiting', 'Nothing pending');
    }

    set('dl-count-done', done.length);
    const monthKey = todayIso().slice(0, 7);
    const thisMonth = done.filter(l => toIsoDate(l.dispatched_on).startsWith(monthKey)).length;
    const noPhoto = done.filter(l => { const e = findEmr(l.emr_no); return !(e && e.photos && e.photos.length); }).length;
    set('dl-note-done', noPhoto ? `${noPhoto} still need a photo` : `${thisMonth} this month`);

    const reserved = getReservedSerials().size;
    set('metric-stock-health', reserved ? `${reserved} reserved on letters` : 'None reserved on letters');
}

function letterItemsHtml(letter) {
    const chips = (letter.items || []).map(i =>
        `<span class="serial-chip" title="${escapeHtml(i.make)} ${escapeHtml(i.model)}">${escapeHtml(i.category)} ${escapeHtml(i.serial_no)}</span>`).join('');
    const others = (letter.other_items || []).map(o =>
        `<span class="serial-chip" style="font-family:var(--font-ui); font-size:.85rem;">${escapeHtml(o.material)} × ${escapeHtml(o.qty)}</span>`).join('');
    return chips + others;
}

function renderLetters() {
    const list = document.getElementById('letter-list');
    if (!list) return;
    const letters = db.dispatch_letters || [];

    const counts = { issued: 0, dispatched: 0, cancelled: 0, all: letters.length };
    letters.forEach(l => { if (counts[l.status] !== undefined) counts[l.status]++; });
    Object.keys(counts).forEach(k => { const el = document.getElementById(`letter-count-${k}`); if (el) el.textContent = counts[k]; });
    document.querySelectorAll('[data-letter-filter]').forEach(b => b.classList.toggle('active', b.getAttribute('data-letter-filter') === letterFilter));

    const q = safeLower((document.getElementById('letter-search') || {}).value);
    let rows = letters.filter(l => letterFilter === 'all' || l.status === letterFilter);
    if (q) {
        rows = rows.filter(l => [l.dl_no, l.destination, l.vendor, l.purpose, l.issued_by, l.emr_no, l.vehicle_no,
            ...(l.items || []).map(i => i.serial_no), ...(l.other_items || []).map(o => o.material)]
            .some(v => safeLower(v).includes(q)));
    }
    rows.sort((a, b) => {
        if (letterFilter === 'issued') return safeDateTimestamp(a.dl_date) - safeDateTimestamp(b.dl_date); // oldest first: do these first
        return safeDateTimestamp(b.dispatched_on || b.dl_date) - safeDateTimestamp(a.dispatched_on || a.dl_date);
    });

    if (!rows.length) {
        const msg = {
            issued: ['No letters waiting at the depot', 'When the office issues a dispatch letter, it shows up here for the depot to dispatch.'],
            dispatched: ['No dispatched letters yet', 'Letters move here once the depot records the dispatch with photos.'],
            cancelled: ['No cancelled letters', ''],
            all: ['No dispatch letters yet', 'Issue the first letter to reserve equipment for an RO.']
        }[letterFilter];
        list.innerHTML = `<div class="empty-state"><strong>${q ? 'No letters match your search' : msg[0]}</strong>${q ? 'Try a serial number, RO name or letter number.' : msg[1]}
            ${!q && (letterFilter === 'issued' || letterFilter === 'all') ? '<br><button class="btn btn-primary" data-open-letter><i class="fa-solid fa-plus"></i> Issue dispatch letter</button>' : ''}</div>`;
        return;
    }

    list.innerHTML = rows.map(l => {
        const e = escapeHtml;
        const age = daysSince(l.dl_date);
        let statusTag = '', side = '', proof = '';
        if (l.status === 'issued') {
            const inStock = assetsInStockCount(l);
            const missing = (l.items || []).length - inStock;
            statusTag = `<span class="tag tag-waiting">Waiting at depot${age ? ` · ${age} day${age === 1 ? '' : 's'}` : ''}</span>`;
            if (missing > 0) statusTag += ` <span class="tag tag-unsv">${missing} not in stock</span>`;
            side = `
                <button class="btn btn-dispatch" data-letter-dispatch="${e(l.dl_no)}"><i class="fa-solid fa-truck-fast"></i> Record dispatch</button>
                <button class="btn btn-outline btn-sm" data-letter-print="${e(l.dl_no)}"><i class="fa-solid fa-print"></i> Print letter</button>
                <button class="btn btn-ghost btn-sm" data-letter-cancel="${e(l.dl_no)}" style="color:var(--red);">Cancel letter</button>`;
        } else if (l.status === 'dispatched') {
            const emr = findEmr(l.emr_no);
            const photos = emr && Array.isArray(emr.photos) ? emr.photos : [];
            statusTag = `<span class="tag tag-new">Dispatched ${e(formatDisplayDate(l.dispatched_on))}</span>` + (photos.length ? '' : ' <span class="tag tag-waiting">Photo missing</span>');
            proof = `<div class="letter-meta">EMR <strong>${e(l.emr_no)}</strong> · vehicle ${e(l.vehicle_no) || '—'} · by ${e(l.dispatched_by) || '—'}</div>
                <div class="letter-proof" data-proof-emr="${e(l.emr_no)}">${photos.slice(0, 6).map(p => `<img class="photo-thumb" data-photo-id="${e(p.id)}" alt="Dispatch photo">`).join('')}</div>`;
            side = `
                <button class="btn btn-outline btn-sm" data-emr-photos="${e(l.emr_no)}"><i class="fa-solid fa-camera"></i> ${photos.length ? `Photos (${photos.length})` : 'Add photo'}</button>
                <button class="btn btn-outline btn-sm" data-emr-print="${e(l.emr_no)}"><i class="fa-solid fa-print"></i> Print EMR</button>
                <button class="btn btn-ghost btn-sm" data-letter-print="${e(l.dl_no)}">Print letter</button>`;
        } else {
            statusTag = `<span class="tag tag-cancel">Cancelled</span>`;
            proof = l.cancel_reason ? `<div class="letter-meta">Reason: ${e(l.cancel_reason)}</div>` : '';
            side = `<button class="btn btn-ghost btn-sm" data-letter-print="${e(l.dl_no)}">Print letter</button>`;
        }
        return `
        <article class="letter-card is-${e(l.status)}" id="letter-${e(l.dl_no).replace(/[^a-zA-Z0-9]/g, '-')}">
            <div class="edge"></div>
            <div class="body">
                <div class="letter-top"><span class="letter-no">${e(l.dl_no)}</span>${statusTag}</div>
                <div class="letter-dest">${e(l.destination)}</div>
                <div class="letter-meta">${e(formatDisplayDate(l.dl_date))} · collect: ${e(l.vendor)}${l.purpose ? ` · ${e(l.purpose)}` : ''} · issued by ${e(l.issued_by)}</div>
                <div class="letter-items">${letterItemsHtml(l)}</div>
                ${l.remarks ? `<div class="letter-meta" style="margin-top:6px;">Instructions: ${e(l.remarks)}</div>` : ''}
                ${proof}
            </div>
            <div class="side">${side}</div>
        </article>`;
    }).join('');

    // Fill photo thumbnails asynchronously
    list.querySelectorAll('img[data-photo-id]').forEach(img => {
        const emr = findEmr(img.closest('[data-proof-emr]').getAttribute('data-proof-emr'));
        const meta = emr && (emr.photos || []).find(p => p.id === img.getAttribute('data-photo-id'));
        photoSrc(meta).then(src => { if (src) img.src = src; else img.alt = 'Photo on another device'; });
    });
}

function gotoLetter(dlNo) {
    const letter = findLetter(dlNo);
    letterFilter = letter ? letter.status : 'all';
    const nav = document.querySelector('.nav-item[data-tab="letters"]');
    if (nav) nav.click();
    const search = document.getElementById('letter-search');
    if (search) search.value = dlNo;
    renderLetters();
}

// ---------- Issue letter (office) ----------
function openLetterModal() {
    letterAssets = [];
    const form = document.getElementById('letter-form');
    if (form) form.reset();
    const date = todayIso();
    document.getElementById('letter-date').value = date;
    document.getElementById('letter-no').value = nextLetterNo(date);
    let issuer = {};
    try { issuer = JSON.parse(localStorage.getItem('hpcl_letter_issuer') || '{}'); } catch (e) { issuer = {}; }
    document.getElementById('letter-issued-by').value = issuer.name || '';
    document.getElementById('letter-designation').value = issuer.designation || '';
    document.getElementById('letter-office').value = issuer.office || 'Belgaum Regional Office';
    document.getElementById('letter-other-lines').innerHTML = '';
    document.getElementById('letter-cat').value = 'DU';
    populateLetterSerialDropdown();
    renderLetterManifest();
    document.getElementById('letter-modal').classList.add('active');
}

function populateLetterSerialDropdown() {
    const cat = document.getElementById('letter-cat').value;
    const sel = document.getElementById('letter-serial-select');
    const list = cat === 'DU' ? db.inventory.pressure_dus : db.inventory.stps;
    const reserved = getReservedSerials();
    const chosen = new Set(letterAssets.map(a => safeLower(a.serial_no)));
    const rows = list
        .filter(i => !chosen.has(safeLower(i.serial_no)) && !reserved.has(safeLower(i.serial_no)))
        .slice()
        .sort((a, b) => safeStr(a.serial_no).localeCompare(safeStr(b.serial_no)));
    sel.innerHTML = `<option value="">Choose… (${rows.length} available)</option>` + rows.map(i => {
        const cond = { N: 'new', U: 'used', S: 'unserviceable' }[safeUpper(i.status)] || safeStr(i.status);
        const spec = cat === 'DU' ? safeStr(i.variant) : `${safeStr(i.hp_type).replace(/\s*HP$/i, '')} HP`;
        return `<option value="${escapeHtml(i.serial_no)}">${escapeHtml(i.serial_no)} · ${escapeHtml(i.make)} ${escapeHtml(i.model)} · ${escapeHtml(spec)} · ${escapeHtml(cond)}</option>`;
    }).join('');
}

function renderLetterManifest() {
    const box = document.getElementById('letter-manifest');
    const tbody = document.getElementById('letter-manifest-tbody');
    const badge = document.getElementById('letter-asset-count');
    badge.textContent = `${letterAssets.length} asset${letterAssets.length === 1 ? '' : 's'}`;
    box.style.display = letterAssets.length ? 'block' : 'none';
    tbody.innerHTML = letterAssets.map((a, i) => `
        <tr>
            <td class="text-muted">${i + 1}</td>
            <td>${escapeHtml(a.category)}</td>
            <td><span class="serial">${escapeHtml(a.serial_no)}</span></td>
            <td>${escapeHtml(a.make)} ${escapeHtml(a.model)} <span class="text-muted">${escapeHtml(a.variant)}</span></td>
            <td>${formatStatusBadge(a.status)}</td>
            <td class="text-right"><button type="button" class="btn-icon btn-delete" data-letter-remove="${escapeHtml(a.serial_no)}" aria-label="Remove ${escapeHtml(a.serial_no)}"><i class="fa-solid fa-xmark"></i></button></td>
        </tr>`).join('');
}

function addLetterAsset() {
    const cat = document.getElementById('letter-cat').value;
    const serial = document.getElementById('letter-serial-select').value;
    if (!serial) { alert('Choose a serial number first.'); return; }
    const list = cat === 'DU' ? db.inventory.pressure_dus : db.inventory.stps;
    const found = list.find(i => safeLower(i.serial_no) === safeLower(serial));
    if (!found) return;
    letterAssets.push(buildEmrAsset(cat, found));
    populateLetterSerialDropdown();
    renderLetterManifest();
}

function addLetterOtherLine(material = '', qty = 1) {
    const wrap = document.getElementById('letter-other-lines');
    const div = document.createElement('div');
    div.className = 'line';
    div.innerHTML = `
        <input type="text" class="form-control" data-other-material placeholder="e.g. Submersible cable 50 m, nozzle 3/4&quot;" value="${escapeHtml(material)}" aria-label="Material">
        <input type="number" class="form-control" data-other-qty min="1" value="${escapeHtml(qty)}" aria-label="Quantity">
        <button type="button" class="btn-icon btn-delete" data-other-remove aria-label="Remove line"><i class="fa-solid fa-xmark"></i></button>`;
    wrap.appendChild(div);
    div.querySelector('[data-other-material]').focus();
}

function readOtherLines() {
    return Array.from(document.querySelectorAll('#letter-other-lines .line')).map(line => ({
        material: safeStr(line.querySelector('[data-other-material]').value),
        qty: parseInt(line.querySelector('[data-other-qty]').value, 10) || 1
    })).filter(l => l.material);
}

function handleLetterSubmit(e) {
    e.preventDefault();
    // include a serial picked but not added
    const pending = document.getElementById('letter-serial-select').value;
    if (pending && !letterAssets.some(a => safeLower(a.serial_no) === safeLower(pending))) addLetterAsset();

    const others = readOtherLines();
    if (!letterAssets.length && !others.length) {
        alert('Add at least one DU, STP or other material line to the letter.');
        return;
    }
    const dlNo = safeStr(document.getElementById('letter-no').value);
    if (findLetter(dlNo)) {
        alert(`Letter no ${dlNo} already exists. Use a different number.`);
        return;
    }
    const reserved = getReservedSerials();
    const clash = letterAssets.filter(a => reserved.has(safeLower(a.serial_no)));
    if (clash.length) {
        alert(`${clash.map(a => a.serial_no).join(', ')} is already reserved on ${reserved.get(safeLower(clash[0].serial_no))}.`);
        return;
    }
    const letter = {
        dl_no: dlNo,
        dl_date: formatDisplayDate(document.getElementById('letter-date').value),
        destination: safeStr(document.getElementById('letter-destination').value),
        vendor: safeStr(document.getElementById('letter-vendor').value),
        purpose: safeStr(document.getElementById('letter-purpose').value),
        issued_by: safeStr(document.getElementById('letter-issued-by').value),
        designation: safeStr(document.getElementById('letter-designation').value),
        office: safeStr(document.getElementById('letter-office').value),
        ref: safeStr(document.getElementById('letter-ref').value),
        remarks: safeStr(document.getElementById('letter-remarks').value),
        items: letterAssets.slice(),
        other_items: others,
        status: 'issued',
        created_at: new Date().toISOString()
    };
    db.dispatch_letters.push(letter);
    try {
        localStorage.setItem('hpcl_letter_issuer', JSON.stringify({ name: letter.issued_by, designation: letter.designation, office: letter.office }));
    } catch (err) { /* not critical */ }
    letterAssets = [];
    saveState();
    refreshAllViews();
    document.getElementById('letter-modal').classList.remove('active');
    showToast(`Dispatch letter ${letter.dl_no} issued. ${letter.items.length} asset(s) reserved.`);
    letterFilter = 'issued';
    printLetter(letter.dl_no);
}

function cancelLetter(dlNo) {
    const l = findLetter(dlNo);
    if (!l || l.status !== 'issued') return;
    const reason = prompt(`Cancel dispatch letter ${l.dl_no}?\n\nReserved assets go back to normal stock. Enter a reason:`, '');
    if (reason === null) return;
    l.status = 'cancelled';
    l.cancel_reason = safeStr(reason);
    l.cancelled_at = new Date().toISOString();
    saveState();
    refreshAllViews();
    showToast(`Letter ${l.dl_no} cancelled.`);
}

// ---------- Record dispatch (depot) ----------
let dispatchingLetter = null;

function nextEmrNo(dateStr) {
    const iso = (dateStr && typeof toIsoDate === 'function' ? toIsoDate(dateStr) : '') || todayIso();
    const prefix = `EMR-${iso.replace(/-/g, '')}-`;
    let seq = 1;
    (db.gate_passes || []).forEach(g => {
        const n = safeStr(g.emr_no);
        if (n.startsWith(prefix)) seq = Math.max(seq, (parseInt(n.slice(prefix.length), 10) || 0) + 1);
    });
    return prefix + String(seq).padStart(3, '0');
}

function openDispatchModal(dlNo) {
    const l = findLetter(dlNo);
    if (!l || l.status !== 'issued') return;
    dispatchingLetter = l;
    const e = escapeHtml;
    const rows = (l.items || []).map((i, idx) => {
        const list = i.category === 'DU' ? db.inventory.pressure_dus : db.inventory.stps;
        const inStock = list.some(x => safeLower(x.serial_no) === safeLower(i.serial_no));
        return `<tr>
            <td class="text-muted">${idx + 1}</td>
            <td>${e(i.category)}</td>
            <td><span class="serial">${e(i.serial_no)}</span></td>
            <td>${e(i.make)} ${e(i.model)}</td>
            <td>${inStock ? '<span class="tag tag-new">In store</span>' : '<span class="tag tag-unsv">Not in store, will be skipped</span>'}</td>
        </tr>`;
    }).join('') + (l.other_items || []).map((o, idx) => `<tr>
            <td class="text-muted">${(l.items || []).length + idx + 1}</td><td>Other</td><td colspan="2">${e(o.material)}</td><td>Qty ${e(o.qty)}</td></tr>`).join('');

    document.getElementById('dispatch-modal-sub').textContent = `${l.dl_no} · ${l.destination}`;
    document.getElementById('dispatch-summary').innerHTML = `
        <div class="callout">Deliver to <strong>${e(l.destination)}</strong>, collected by <strong>${e(l.vendor)}</strong>${l.purpose ? ` for ${e(l.purpose)}` : ''}.
        ${l.remarks ? `<br>Office instructions: ${e(l.remarks)}` : ''}</div>
        <div class="manifest" style="margin:0 0 16px;"><table>
            <thead><tr><th>#</th><th>Type</th><th>Serial no</th><th>Make / model</th><th>Check</th></tr></thead>
            <tbody>${rows}</tbody>
        </table></div>`;

    const date = todayIso();
    document.getElementById('dispatch-out-date').value = date;
    document.getElementById('dispatch-emr-no').value = nextEmrNo(date);
    ['dispatch-vehicle', 'dispatch-driver', 'dispatch-gp-no', 'dispatch-remarks'].forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById('dispatch-by').value = localStorage.getItem('hpcl_dispatch_by') || '';
    resetPhotoCapture('dispatch');
    document.getElementById('dispatch-modal').classList.add('active');
}

async function handleDispatchSubmit(e) {
    e.preventDefault();
    const l = dispatchingLetter;
    if (!l) return;
    if (captureCount('dispatch') === 0) {
        const cap = photoCaptures.dispatch;
        if (cap) cap.root.classList.add('required-missing');
        alert('Take at least one photo of the loaded vehicle before confirming the dispatch.');
        return;
    }
    const outDate = formatDisplayDate(document.getElementById('dispatch-out-date').value);
    const dispatchedBy = safeStr(document.getElementById('dispatch-by').value);
    let result;
    try {
        result = createEmr({
            assets: l.items || [],
            otherItems: l.other_items || [],
            emrNumber: document.getElementById('dispatch-emr-no').value,
            outDate,
            vendor: l.vendor,
            destination: l.destination,
            gatePassNo: document.getElementById('dispatch-gp-no').value,
            vehicleNo: document.getElementById('dispatch-vehicle').value,
            driver: document.getElementById('dispatch-driver').value,
            remarks: document.getElementById('dispatch-remarks').value || l.remarks,
            dlNo: l.dl_no,
            dispatchedBy
        });
    } catch (err) {
        alert(err.message);
        return;
    }
    await commitCapturedPhotos('dispatch', result.emr);

    l.status = 'dispatched';
    l.emr_no = result.emr.emr_no;
    l.dispatched_on = outDate;
    l.dispatched_by = dispatchedBy;
    l.vehicle_no = result.emr.vehicle_no;
    l.driver = result.emr.driver;
    l.gate_pass_no = result.emr.gate_pass_no;
    if (result.skipped.length) l.skipped_serials = result.skipped;
    try { localStorage.setItem('hpcl_dispatch_by', dispatchedBy); } catch (err) { /* ignore */ }

    letterFilter = 'dispatched';
    saveState();
    refreshAllViews();
    document.getElementById('dispatch-modal').classList.remove('active');
    dispatchingLetter = null;
    showToast(`Dispatched against ${l.dl_no}. EMR ${l.emr_no} created with ${result.emr.photos.length} photo(s).`
        + (result.skipped.length ? ` Skipped (not in store): ${result.skipped.join(', ')}.` : ''));
    uploadPendingPhotos();
}

// ---------- Photos for an EMR ----------
let photosModalEmr = null;

async function openPhotosModal(emrNo) {
    const emr = findEmr(emrNo);
    if (!emr) return;
    photosModalEmr = emr;
    if (!Array.isArray(emr.photos)) emr.photos = [];
    document.getElementById('photos-modal-title').textContent = `Dispatch photos · ${emr.emr_no}`;
    document.getElementById('photos-modal-sub').textContent = `${emr.destination || ''}${emr.vehicle_no ? ' · ' + emr.vehicle_no : ''}${emr.out_date ? ' · ' + formatDisplayDate(emr.out_date) : ''}`;
    resetPhotoCapture('extra');
    await renderPhotosGallery();
    document.getElementById('photos-modal').classList.add('active');
}

async function renderPhotosGallery() {
    const emr = photosModalEmr;
    const gallery = document.getElementById('photos-gallery');
    if (!emr || !gallery) return;
    if (!emr.photos.length) {
        gallery.innerHTML = '<div class="empty-state" style="grid-column:1/-1;"><strong>No photos yet</strong>Add a photo of the loaded vehicle and the serial plates below.</div>';
        return;
    }
    gallery.innerHTML = emr.photos.map(p => {
        const t = p.taken_at ? formatStampTime(new Date(p.taken_at)) : '';
        const where = p.drive_id ? `<a href="${escapeHtml(p.drive_url)}" target="_blank" rel="noopener">Drive</a>` : 'This device only';
        return `<figure>
            <img data-gallery-id="${escapeHtml(p.id)}" alt="Dispatch photo ${escapeHtml(t)}">
            <figcaption><span>${escapeHtml(t)} · ${where}</span>
            <button type="button" class="btn-icon btn-delete" data-photo-delete="${escapeHtml(p.id)}" title="Delete photo" aria-label="Delete photo" style="width:26px;height:26px;"><i class="fa-solid fa-trash-can"></i></button></figcaption>
        </figure>`;
    }).join('');
    for (const img of gallery.querySelectorAll('img[data-gallery-id]')) {
        const meta = emr.photos.find(p => p.id === img.getAttribute('data-gallery-id'));
        const src = await photoSrc(meta);
        if (src) img.src = src; else img.alt = 'This photo is on another device and not uploaded yet';
    }
}

async function savePhotosModal() {
    const emr = photosModalEmr;
    if (!emr) return;
    if (!captureCount('extra')) {
        document.getElementById('photos-modal').classList.remove('active');
        return;
    }
    const metas = await commitCapturedPhotos('extra', emr);
    saveState();
    refreshAllViews();
    await renderPhotosGallery();
    showToast(`${metas.length} photo(s) added to ${emr.emr_no}.`);
    uploadPendingPhotos();
}

function deletePhoto(photoId) {
    const emr = photosModalEmr;
    if (!emr) return;
    const meta = emr.photos.find(p => p.id === photoId);
    if (!meta) return;
    if (!confirm(`Delete this photo from ${emr.emr_no}?${meta.drive_id ? '\n\nThe copy already in Google Drive is not deleted.' : ''}`)) return;
    emr.photos = emr.photos.filter(p => p.id !== photoId);
    deletePhotosLocal([photoId]);
    saveState();
    refreshAllViews();
    renderPhotosGallery();
}

// ---------- Printable dispatch letter ----------
function printLetter(dlNo) {
    const l = findLetter(dlNo);
    if (!l) return;
    const e = escapeHtml;
    const modal = document.getElementById('gate-pass-print-modal');
    const body = document.getElementById('printable-voucher-body');
    document.getElementById('print-modal-title').textContent = 'Dispatch letter';
    const cell = 'padding:8px 10px; border:1px solid #B7C4D2;';
    const cond = s => ({ N: 'New', U: 'Used', S: 'Unserviceable' }[safeUpper(s)] || safeStr(s) || '—');
    const rows = (l.items || []).map((i, idx) => `
        <tr>
            <td style="${cell} text-align:center;">${idx + 1}</td>
            <td style="${cell}">${i.category === 'DU' ? 'Dispensing unit' : 'Submersible pump'}<br><span style="color:#3D5470; font-size:12px;">${e(i.variant || '')}</span></td>
            <td style="${cell}">${e(i.make)} ${e(i.model)}</td>
            <td style="${cell} font-family:'Barlow Condensed',sans-serif; font-weight:600; font-size:15px;">${e(i.serial_no)}</td>
            <td style="${cell}">${e(cond(i.status))}</td>
            <td style="${cell} text-align:center;">1</td>
        </tr>`).join('') + (l.other_items || []).map((o, idx) => `
        <tr>
            <td style="${cell} text-align:center;">${(l.items || []).length + idx + 1}</td>
            <td style="${cell}" colspan="4">${e(o.material)}</td>
            <td style="${cell} text-align:center;">${e(o.qty)}</td>
        </tr>`).join('');
    const disp = l.status === 'dispatched';
    body.innerHTML = `
    <div class="printable-letter-sheet" style="font-family:'Barlow',Arial,sans-serif; color:#0F2A4A; font-size:13.5px; line-height:1.5; display:flex; flex-direction:column;">
        <div class="print-avoid-break print-masthead" style="border-bottom:3px solid #0F2A4A; padding-bottom:12px;">
            <div class="print-brand-row">
                <img src="HPCL.png" alt="HPCL">
                <div>
                    <div style="font-family:'Noto Sans Devanagari',sans-serif; font-weight:700; color:#D2232A; font-size:15px;">हिन्दुस्तान पेट्रोलियम कॉर्पोरेशन लिमिटेड</div>
                    <div class="print-brand-title-en" style="font-family:'Barlow Condensed',sans-serif; font-weight:700; font-size:22px; line-height:1.1;">HINDUSTAN PETROLEUM CORPORATION LIMITED</div>
                    <div style="color:#3D5470;">${e(l.office || 'Belgaum Regional Office')}</div>
                </div>
            </div>
            <div class="print-masthead-right" style="font-size:13px;">
                <div><strong>Ref:</strong> ${e(l.dl_no)}</div>
                ${l.ref ? `<div><strong>File:</strong> ${e(l.ref)}</div>` : ''}
                <div><strong>Date:</strong> ${e(formatDisplayDate(l.dl_date))}</div>
                ${l.status === 'cancelled' ? '<div style="color:#D2232A; font-weight:700; margin-top:4px;">CANCELLED</div>' : ''}
            </div>
        </div>

        <p style="margin:18px 0 0;">To,<br>The Depot In-charge,<br>Desur IRD Depot, Belgaum</p>
        <p style="margin:14px 0;"><strong>Subject: Release of equipment for ${e(l.destination)}</strong></p>
        <p style="margin:0 0 12px;">Please release the following equipment from Desur IRD stock to <strong>${e(l.vendor)}</strong>${l.purpose ? ` for <strong>${e(l.purpose)}</strong>` : ''} at <strong>${e(l.destination)}</strong>.
        At the time of dispatch, issue the EMR, record the vehicle number and add photographs of the loaded material in the stores register.</p>

        <div class="print-table-wrap">
            <table style="width:100%; border-collapse:collapse; font-size:13px; margin:10px 0 14px;">
                <thead><tr style="background:#0F2A4A; color:#fff;">
                    <th style="${cell} border-color:#0F2A4A; width:36px;">#</th>
                    <th style="${cell} border-color:#0F2A4A; text-align:left;">Equipment</th>
                    <th style="${cell} border-color:#0F2A4A; text-align:left;">Make &amp; model</th>
                    <th style="${cell} border-color:#0F2A4A; text-align:left;">Serial no</th>
                    <th style="${cell} border-color:#0F2A4A; text-align:left;">Condition</th>
                    <th style="${cell} border-color:#0F2A4A; width:50px;">Qty</th>
                </tr></thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
        ${l.remarks ? `<p style="margin:0 0 12px;"><strong>Instructions:</strong> ${e(l.remarks)}</p>` : ''}

        <div class="print-bottom-block" style="margin-top:auto; padding-top:24px;">
            <div style="display:flex; justify-content:flex-end;">
                <div style="text-align:left; min-width:220px;">
                    <div style="border-top:1px solid #0F2A4A; padding-top:6px;"><strong>${e(l.issued_by)}</strong></div>
                    <div>${e(l.designation)}</div>
                    <div style="color:#3D5470;">${e(l.office)}</div>
                </div>
            </div>

            <div style="margin-top:22px; border:1.5px dashed #6B7F96; padding:12px 14px; font-size:13px;">
                <div style="font-family:'Barlow Condensed',sans-serif; font-weight:700; font-size:16px; margin-bottom:6px;">Depot endorsement</div>
                <div class="print-endorse-grid">
                    <div>Dispatched on: <strong>${disp ? e(formatDisplayDate(l.dispatched_on)) : '____________'}</strong></div>
                    <div>EMR no: <strong>${disp ? e(l.emr_no) : '____________'}</strong></div>
                    <div>Vehicle no: <strong>${disp ? e(l.vehicle_no) : '____________'}</strong></div>
                    <div>Dispatched by: <strong>${disp ? e(l.dispatched_by) : '____________'}</strong></div>
                </div>
            </div>
            <div class="print-footer-row" style="margin-top:12px; border-top:1px solid #D5DDE6; padding-top:8px; font-size:11px; color:#6B7F96;">
                <span>HPCL Desur IRD Depot · Belgaum Regional Office (RRO)</span>
                <span>Generated from Stores Register</span>
            </div>
        </div>
    </div>`;
    modal.classList.add('active');
}

// ---------- Datalists (vendor / RO suggestions) ----------
function refreshDatalists() {
    const fill = (id, values) => {
        const el = document.getElementById(id);
        if (!el) return;
        const uniq = Array.from(new Set(values.map(v => safeStr(v)).filter(Boolean))).sort((a, b) => a.localeCompare(b)).slice(0, 300);
        el.innerHTML = uniq.map(v => `<option value="${escapeHtml(v)}"></option>`).join('');
    };
    const src = [...(db.gate_passes || []), ...(db.du_outword || []), ...(db.stp_outword || []), ...(db.dispatch_letters || [])];
    fill('vendor-suggestions', src.map(r => r.vendor));
    fill('destination-suggestions', src.map(r => r.destination));
}

// --------------------------------------------------------------------------
// Wiring
// --------------------------------------------------------------------------
(function wrapRefresh() {
    const base = refreshAllViews;
    refreshAllViews = function () {
        base();
        renderLetters();
        renderDispatchLine();
        refreshDatalists();
        updatePendingPhotoCount();
    };
})();

document.addEventListener('DOMContentLoaded', () => {
    initPhotoCaptures();

    document.addEventListener('click', (e) => {
        const t = e.target;
        if (t.closest('[data-open-letter]')) { e.preventDefault(); openLetterModal(); return; }
        const goto = t.closest('[data-goto-letters]');
        if (goto) {
            e.preventDefault();
            letterFilter = goto.getAttribute('data-goto-letters');
            const nav = document.querySelector('.nav-item[data-tab="letters"]');
            if (nav) nav.click();
            return;
        }
        const f = t.closest('[data-letter-filter]');
        if (f) { letterFilter = f.getAttribute('data-letter-filter'); renderLetters(); return; }
        const d = t.closest('[data-letter-dispatch]');
        if (d) { openDispatchModal(d.getAttribute('data-letter-dispatch')); return; }
        const pr = t.closest('[data-letter-print]');
        if (pr) { printLetter(pr.getAttribute('data-letter-print')); return; }
        const c = t.closest('[data-letter-cancel]');
        if (c) { cancelLetter(c.getAttribute('data-letter-cancel')); return; }
        const ph = t.closest('[data-emr-photos]');
        if (ph) { openPhotosModal(ph.getAttribute('data-emr-photos')); return; }
        const pe = t.closest('[data-emr-print]');
        if (pe) { openGatePassPrintModal(pe.getAttribute('data-emr-print')); return; }
        const rm = t.closest('[data-letter-remove]');
        if (rm) {
            letterAssets = letterAssets.filter(a => safeLower(a.serial_no) !== safeLower(rm.getAttribute('data-letter-remove')));
            populateLetterSerialDropdown();
            renderLetterManifest();
            return;
        }
        const ol = t.closest('[data-other-remove]');
        if (ol) { ol.closest('.line').remove(); return; }
        const del = t.closest('[data-photo-delete]');
        if (del) { deletePhoto(del.getAttribute('data-photo-delete')); return; }
        const zoom = t.closest('img.photo-thumb, img[data-gallery-id]');
        if (zoom && zoom.src) { openLightbox(zoom.src, zoom.alt); }
    });

    document.getElementById('letter-cat').addEventListener('change', populateLetterSerialDropdown);
    document.getElementById('letter-add-asset-btn').addEventListener('click', addLetterAsset);
    document.getElementById('letter-add-line-btn').addEventListener('click', () => addLetterOtherLine());
    document.getElementById('letter-form').addEventListener('submit', handleLetterSubmit);
    document.getElementById('letter-date').addEventListener('change', (e) => {
        const noInput = document.getElementById('letter-no');
        if (/^DL\/DESUR\//.test(noInput.value)) noInput.value = nextLetterNo(e.target.value);
    });
    document.getElementById('dispatch-form').addEventListener('submit', handleDispatchSubmit);
    document.getElementById('photos-save-btn').addEventListener('click', savePhotosModal);
    document.getElementById('letter-search').addEventListener('input', renderLetters);
    document.getElementById('btn-export-letters').addEventListener('click', () => exportCollectionToCsv('letters'));
    document.getElementById('btn-sync-photos').addEventListener('click', () => uploadPendingPhotos({ manual: true }));
    document.getElementById('lightbox-close').addEventListener('click', closeLightbox);
    document.getElementById('lightbox').addEventListener('click', (e) => { if (e.target.id === 'lightbox') closeLightbox(); });
    window.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeLightbox(); });
    window.addEventListener('online', () => setTimeout(() => uploadPendingPhotos(), 3000));

    // Open the tab named in the URL (#letters etc.) once data has loaded
    setTimeout(() => {
        const hash = (location.hash || '').replace('#', '');
        if (hash && hash !== 'dashboard') {
            const nav = document.querySelector(`.nav-item[data-tab="${hash}"]`);
            if (nav) nav.click();
        }
        uploadPendingPhotos();
    }, 600);
});
