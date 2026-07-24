// ============================================================
// Printposs KPI Dashboard - Unified Single-Page Application
// ============================================================

// === Constants ===
const FACTORY_MAP = { 'hoson': 'Factory A', 'jz': 'Factory B', 'monkeyprint': 'Factory C' };
const HOSON_MASKED = 'Factory A';

function maskFactory(name) {
    if (!name) return '-';
    const lower = name.trim().toLowerCase();
    return FACTORY_MAP[lower] || ('Factory ' + name.trim());
}

// === DOM Elements ===
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadStatus = document.getElementById('upload-status');
const navItems = document.querySelectorAll('.nav-item');
const tabPanes = document.querySelectorAll('.tab-pane');
const hiddenPreUpload = document.querySelectorAll('.hidden-pre-upload');
const pageTitle = document.getElementById('page-title');
const pageSubtitle = document.getElementById('page-subtitle');
const btnReset = document.getElementById('btn-reset');

// Filters
const sellerFilter = document.getElementById('seller-filter');
const yearFilter = document.getElementById('year-filter');
const monthFilter = document.getElementById('month-filter');
const startDateInput = document.getElementById('start-date');
const endDateInput = document.getElementById('end-date');

// Global State
let transformedData = [];
let chartInstances = {};

// Chart.js defaults for dark mode
Chart.defaults.color = '#94a3b8';
Chart.defaults.borderColor = '#334155';

// === Center Text Plugin for Doughnut Charts ===
const centerTextPlugin = {
    id: 'centerText',
    beforeDraw: function(chart) {
        if (chart.config.type !== 'doughnut') return;
        const ct = chart.config.options.plugins && chart.config.options.plugins.centerText;
        if (!ct) return;
        const { ctx, chartArea: { left, top, width, height } } = chart;
        const cx = left + width / 2;
        const cy = top + height / 2;
        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = 'bold 22px Outfit, sans-serif';
        ctx.fillStyle = '#f8fafc';
        ctx.fillText(ct.text || '', cx, cy - 8);
        ctx.font = '11px Inter, sans-serif';
        ctx.fillStyle = '#94a3b8';
        ctx.fillText(ct.subtext || '', cx, cy + 14);
        ctx.restore();
    }
};
Chart.register(centerTextPlugin);

// === Drag & Drop ===
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); }, false);
});
['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, () => dropZone.classList.add('dragover'), false);
});
['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, () => dropZone.classList.remove('dragover'), false);
});
dropZone.addEventListener('drop', e => handleFiles(e.dataTransfer.files), false);
fileInput.addEventListener('change', function() {
    if (this.files.length) handleFiles(this.files);
});

// === Reset Button ===
btnReset.addEventListener('click', () => {
    transformedData = [];
    sellerFilter.value = 'all';
    yearFilter.value = 'all';
    monthFilter.value = 'all';
    startDateInput.value = '';
    endDateInput.value = '';
    hiddenPreUpload.forEach(el => el.classList.add('hidden'));
    switchTab('tab-upload');
    fileInput.value = '';
});

window.addEventListener('DOMContentLoaded', () => {
    fetch('file-list.json')
        .then(res => {
            if (!res.ok) throw new Error("No file-list.json");
            return res.json();
        })
        .then(files => {
            if (files && files.length > 0) {
                hiddenPreUpload.forEach(el => el.classList.remove('hidden'));
                switchTab('tab-dashboard', document.querySelector('[data-tab="dashboard"]'));
                uploadStatus.classList.remove('hidden');
                
                let allData = [];
                let done = 0;
                const finish = () => {
                    uploadStatus.classList.add('hidden');
                    const deduped = deduplicateRawData(allData);
                    const newOrders = transformData(deduped);
                    mergeNewOrders(newOrders);
                    populateSellerFilter();
                    populateYearFilter();
                    updateSidebarCount();
                    renderDashboard();
                };

                files.forEach(file => {
                    fetch('data/' + file)
                        .then(r => r.blob())
                        .then(blob => {
                            // processFile requires access to allData, done, finish from this scope
                            // So we need to re-implement the processFile logic here briefly, or move processFile out.
                            const reader = new FileReader();
                            reader.onload = e => {
                                const bytes = new Uint8Array(e.target.result);
                                let encoding = 'windows-1252'; // Default for VN Excel
                                if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
                                    encoding = 'utf-8'; // Excel CSV UTF-8 always has BOM
                                }
                                Papa.parse(blob, {
                                    header: true, skipEmptyLines: true, worker: true,
                                    encoding: encoding,
                                    complete: r => { 
                                        if (r.data.length > 0) {
                                            const keyMap = {};
                                            let needsRename = false;
                                            for (let k in r.data[0]) {
                                                const clean = k.trim().toLowerCase().replace(/^\uFEFF/, '');
                                                keyMap[k] = clean;
                                                if (k !== clean) needsRename = true;
                                            }
                                            if (needsRename) {
                                                for (let i = 0; i < r.data.length; i++) {
                                                    const oldObj = r.data[i];
                                                    const newObj = {};
                                                    for (let k in oldObj) newObj[keyMap[k]] = oldObj[k];
                                                    r.data[i] = newObj;
                                                }
                                            }
                                        }
                                        allData = allData.concat(r.data); 
                                        done++;
                                        if (done === files.length) finish();
                                    },
                                    error: () => { done++; if (done === files.length) finish(); }
                                });
                            };
                            const chunk = blob.slice(0, 1024 * 1024);
                            reader.readAsArrayBuffer(chunk);
                        })
                        .catch(err => {
                            console.error('Error fetching file', file, err);
                            done++;
                            if (done === files.length) finish();
                        });
                });
            }
        })
        .catch(e => {
            // No file-list.json found, fallback to preloaded data if any
            if (typeof PRELOADED_ROWS !== 'undefined' && Array.isArray(PRELOADED_ROWS) && PRELOADED_ROWS.length > 0) {
                transformedData = PRELOADED_ROWS.map(row => {
                    const obj = {};
                    PRELOADED_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
                    obj.created_obj = obj.created_time ? new Date(obj.created_time * 1000) : null;
                    obj.paid_obj = obj.paid_time ? new Date(obj.paid_time * 1000) : null;
                    obj.delivered_obj = obj.delivered_time ? new Date(obj.delivered_time * 1000) : null;
                    obj.has_ticket = obj.has_ticket === 1;
                    const sh = obj.shipped_at !== '-' ? parseCustomDate(obj.shipped_at) : null;
                    if (sh && obj.delivered_obj) obj.ship_wd = getWorkingDays(sh, obj.delivered_obj);
                    return obj;
                });
                hiddenPreUpload.forEach(el => el.classList.remove('hidden'));
                populateSellerFilter();
                populateYearFilter();
                updateSidebarCount();
                renderDashboard();
                switchTab('tab-dashboard', document.querySelector('[data-tab="dashboard"]'));
            }
        });
});

// === Filter Helpers ===
function populateSellerFilter() {
    const select = sellerFilter;
    const cur = select.value;
    const sellers = new Set();
    transformedData.forEach(d => { if (d.seller && d.seller !== 'Unknown') sellers.add(d.seller); });
    const sorted = Array.from(sellers).sort();
    select.innerHTML = '<option value="all">All Sellers</option>';
    sorted.forEach(s => { select.innerHTML += '<option value="' + s + '">' + s + '</option>'; });
    select.value = Array.from(select.options).some(o => o.value === cur) ? cur : 'all';
}

function populateYearFilter() {
    const select = yearFilter;
    const cur = select.value;
    const years = new Set();
    transformedData.forEach(d => { if (d.created_obj) years.add(d.created_obj.getFullYear()); });
    const sorted = Array.from(years).sort((a, b) => b - a);
    select.innerHTML = '<option value="all">All Years</option>';
    sorted.forEach(y => { select.innerHTML += '<option value="' + y + '">' + y + '</option>'; });
    select.value = Array.from(select.options).some(o => o.value === cur) ? cur : 'all';
}

function getFilteredData() {
    let data = transformedData;
    const sv = sellerFilter.value;
    if (sv !== 'all') data = data.filter(d => d.seller === sv);
    const yv = yearFilter.value;
    if (yv !== 'all') data = data.filter(d => d.created_obj && d.created_obj.getFullYear().toString() === yv);
    const mv = monthFilter.value;
    if (mv !== 'all') data = data.filter(d => d.created_obj && (d.created_obj.getMonth() + 1).toString() === mv);
    if (startDateInput.value) {
        const s = new Date(startDateInput.value); s.setHours(0, 0, 0, 0);
        data = data.filter(d => d.created_obj && d.created_obj >= s);
    }
    if (endDateInput.value) {
        const e = new Date(endDateInput.value); e.setHours(23, 59, 59, 999);
        data = data.filter(d => d.created_obj && d.created_obj <= e);
    }
    return data;
}

function updateSidebarCount() {
    const el = document.getElementById('sidebar-order-count');
    if (el) el.innerText = formatNumber(transformedData.length) + ' orders';
}

// === Filter Event Listeners ===
[sellerFilter, yearFilter, monthFilter, startDateInput, endDateInput].forEach(el => {
    if (el) {
        el.addEventListener('change', () => renderDashboard());
        if (el.type === 'date') el.addEventListener('input', () => renderDashboard());
    }
});

// === Tab Navigation ===
navItems.forEach(item => {
    item.addEventListener('click', e => {
        e.preventDefault();
        if (item.classList.contains('hidden')) return;
        const tab = item.getAttribute('data-tab');
        if (tab) switchTab('tab-' + tab, item);
    });
});

function switchTab(tabId, navEl) {
    tabPanes.forEach(p => { p.classList.add('hidden'); p.classList.remove('active'); });
    const target = document.getElementById(tabId);
    if (target) { target.classList.remove('hidden'); target.classList.add('active'); }
    navItems.forEach(n => n.classList.remove('active'));
    if (navEl) navEl.classList.add('active');

    if (tabId === 'tab-dashboard') {
        pageTitle.innerText = 'KPI Dashboard';
        pageSubtitle.innerText = 'Comprehensive order lifecycle analytics';
        renderDashboard();
    } else if (tabId === 'tab-upload') {
        pageTitle.innerText = 'Upload Data';
        pageSubtitle.innerText = 'Upload CSV files to analyze order lifecycle data.';
    }
}

// === Data Processing ===
function handleFiles(files) {
    const csvFiles = Array.from(files).filter(f => f.name.endsWith('.csv'));
    if (!csvFiles.length) { alert('No CSV files found.'); return; }
    processMultipleCSVs(csvFiles);
}

function processMultipleCSVs(files) {
    uploadStatus.classList.remove('hidden');
    let allData = [];
    let done = 0;
    const finish = () => {
        uploadStatus.classList.add('hidden');
        const deduped = deduplicateRawData(allData);
        const newOrders = transformData(deduped);
        mergeNewOrders(newOrders);
        hiddenPreUpload.forEach(el => el.classList.remove('hidden'));
        populateSellerFilter();
        populateYearFilter();
        updateSidebarCount();
        renderDashboard();
        switchTab('tab-dashboard', document.querySelector('[data-tab="dashboard"]'));
    };
    const processFile = (file, callback) => {
        const reader = new FileReader();
        reader.onload = e => {
            const bytes = new Uint8Array(e.target.result);
            let encoding = 'windows-1252'; // Default for VN Excel
            if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
                encoding = 'utf-8';
            }
            Papa.parse(file, {
                header: true, skipEmptyLines: true, worker: true,
                encoding: encoding,
                complete: r => { 
                    if (r.data.length > 0) {
                        const keyMap = {};
                        let needsRename = false;
                        for (let k in r.data[0]) {
                            const clean = k.trim().toLowerCase().replace(/^\uFEFF/, '');
                            keyMap[k] = clean;
                            if (k !== clean) needsRename = true;
                        }
                        if (needsRename) {
                            for (let i = 0; i < r.data.length; i++) {
                                const oldObj = r.data[i];
                                const newObj = {};
                                for (let k in oldObj) newObj[keyMap[k]] = oldObj[k];
                                r.data[i] = newObj;
                            }
                        }
                    }
                    allData = allData.concat(r.data); 
                    callback(); 
                },
                error: () => { callback(); }
            });
        };
        // Only read the first 1MB to detect encoding to avoid freezing UI on large files
        const chunk = file.slice(0, 1024 * 1024);
        reader.readAsArrayBuffer(chunk);
    };

    files.forEach(file => {
        processFile(file, () => {
            done++;
            if (done === files.length) finish();
        });
    });
}

function deduplicateRawData(arr) {
    const seen = new Set();
    return arr.filter(row => {
        const id = row.order_number || row['order number'];
        if (!id) return false;
        const eventName = row.event_name || row['event name'] || '';
        const updateOrder = row.update_order || row['update order'] || '';
        const key = id.trim() + '|' + eventName.trim() + '|' + updateOrder.trim();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

// === Date Utilities ===
function parseCustomDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    let p = new Date(dateStr.trim());
    if (!isNaN(p.getTime())) return p;
    try {
        const [dp, tp] = dateStr.trim().split(' ');
        if (!dp) return null;
        const [m, d, y] = dp.split('/');
        const hm = tp ? tp.split(':') : ['0', '0'];
        p = new Date(y, m - 1, d, hm[0], hm[1] || '0');
        return isNaN(p.getTime()) ? null : p;
    } catch (e) { return null; }
}

function formatDateTime(date) {
    if (!date) return '-';
    const yy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mi = String(date.getMinutes()).padStart(2, '0');
    return yy + '-' + mm + '-' + dd + ' ' + hh + ':' + mi;
}

function formatDateOnly(date) {
    if (!date) return '-';
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
}

function formatNumber(n) {
    return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function getWorkingDays(startDate, endDate) {
    if (!startDate || !endDate || startDate >= endDate) return 0;
    let days = 0;
    let current = new Date(startDate);
    let totalMs = endDate - startDate;
    let fullDays = Math.floor(totalMs / 86400000);
    let fractionalDay = (totalMs % 86400000) / 86400000;
    
    for (let i = 0; i < fullDays; i++) {
        let dayOfWeek = current.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) days++;
        current.setDate(current.getDate() + 1);
    }
    
    let dayOfWeek = current.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) days += fractionalDay;
    return days;
}

// === Transform Uploaded CSV Data ===
function transformData(rawData) {
    const orderMap = new Map();
    rawData.forEach(row => {
        const id = row.order_number || row['order number'];
        if (!id) return;
        if (!orderMap.has(id)) {
            orderMap.set(id, { order_number: id, cohort_date_obj: null, status: null, factory_raw: null, seller: null, has_ticket: false, events: {} });
        }
        const order = orderMap.get(id);
        const orderStatus = row.order_status || row['order status'];
        if (orderStatus && !order.status) order.status = orderStatus;
        const fp = row.fulfillment_partner || row['fulfillment partner'];
        if (fp && !order.factory_raw) order.factory_raw = fp.trim();
        const sv = row.seller_name || row.seller || row['seller name'];
        if (sv && !order.seller) order.seller = sv.trim();

        // Ticket detection
        const tv = row.ticket;
        if (tv && tv.trim().toLowerCase() === 'x') order.has_ticket = true;

        const eventName = row.event_name || row['event name'];
        const updateDate = parseCustomDate(row.update_order || row['update order']);
        if (eventName) order.events[eventName] = updateDate;
        const createdOrderDate = row.created_order || row['created order'];
        if (createdOrderDate && !order.cohort_date_obj) order.cohort_date_obj = parseCustomDate(createdOrderDate);
    });

    return Array.from(orderMap.values()).map(order => {
        const e = order.events;
        const baseDate = order.cohort_date_obj || e['order_created'] || new Date();
        const isCan = order.status && order.status.toLowerCase().includes('cancel');
        const isH = order.factory_raw && order.factory_raw.toLowerCase().includes('hoson');

        const cr = e['order_created'], pa = e['order_paid'], di = e['order_dispatched'], sh = e['order_shipped'], de = e['order_delivered'];
        let sl_lt = null, op_lt = null, pr_lt = null, lg_lt = null;
        let sl_sla = 'Pending', op_sla = 'Pending', pr_sla = 'Pending', lg_sla = 'Pending';

        if (isCan) {
            sl_sla = op_sla = pr_sla = lg_sla = 'Cancelled';
        } else {
            if (cr && pa) { sl_lt = (pa - cr) / 36e5; sl_sla = sl_lt <= 24 ? 'Achieved' : 'Overdue'; } else if (cr) sl_sla = 'In Progress';
            if (pa && di) { op_lt = (di - pa) / 36e5; op_sla = op_lt <= 1 ? 'Achieved' : 'Overdue'; } else if (pa) op_sla = 'In Progress';
            const ps = isH ? pa : di;
            if (ps && sh) { pr_lt = (sh - ps) / 36e5; pr_sla = pr_lt <= 24 ? 'Achieved' : 'Overdue'; } else if (ps) pr_sla = 'In Progress';
            if (sh && de) { lg_lt = (de - sh) / 36e5; lg_sla = lg_lt <= 168 ? 'Achieved' : 'Overdue'; } else if (sh) lg_sla = 'In Progress';
        }

        return {
            order_number: order.order_number, status: order.status || '-',
            factory: maskFactory(order.factory_raw), seller: order.seller || 'Unknown',
            cohort_date: formatDateOnly(baseDate),
            created_at: formatDateTime(cr), paid_at: formatDateTime(pa),
            dispatched_at: formatDateTime(di), shipped_at: formatDateTime(sh), delivered_at: formatDateTime(de),
            sl_lt, sl_sla, op_lt, op_sla, pr_lt, pr_sla, lg_lt, lg_sla,
            created_obj: cr, paid_obj: pa, delivered_obj: de,
            has_ticket: order.has_ticket
        };
    });
}

// === Merge Uploaded Orders with Preloaded ===
function mergeNewOrders(newOrders) {
    const map = new Map();
    transformedData.forEach(o => map.set(o.order_number, o));

    newOrders.forEach(n => {
        if (map.has(n.order_number)) {
            const old = map.get(n.order_number);
            const m = { ...old };
            if (n.status !== '-') m.status = n.status;
            if (n.factory !== '-') m.factory = n.factory;
            if (n.seller !== 'Unknown') m.seller = n.seller;
            if (n.has_ticket) m.has_ticket = true;

            ['created_at', 'paid_at', 'dispatched_at', 'shipped_at', 'delivered_at'].forEach(f => { if (n[f] !== '-') m[f] = n[f]; });
            ['created_obj', 'paid_obj', 'delivered_obj'].forEach(f => { if (n[f]) m[f] = n[f]; });

            // Recalculate SLA
            const isH = m.factory === HOSON_MASKED;
            const isCan = m.status.toLowerCase().includes('cancel');
            const cr = m.created_obj, pa = m.paid_obj;
            const di = m.dispatched_at !== '-' ? parseCustomDate(m.dispatched_at) : null;
            const sh = m.shipped_at !== '-' ? parseCustomDate(m.shipped_at) : null;
            const de = m.delivered_obj;

            if (isCan) {
                m.sl_sla = m.op_sla = m.pr_sla = m.lg_sla = 'Cancelled';
                m.sl_lt = m.op_lt = m.pr_lt = m.lg_lt = null;
            } else {
                if (cr && pa) { m.sl_lt = (pa - cr) / 36e5; m.sl_sla = m.sl_lt <= 24 ? 'Achieved' : 'Overdue'; }
                if (pa && di) { m.op_lt = (di - pa) / 36e5; m.op_sla = m.op_lt <= 1 ? 'Achieved' : 'Overdue'; }
                const ps = isH ? pa : di;
                if (ps && sh) { m.pr_lt = (sh - ps) / 36e5; m.pr_sla = m.pr_lt <= 24 ? 'Achieved' : 'Overdue'; }
                if (sh && de) { m.lg_lt = (de - sh) / 36e5; m.lg_sla = m.lg_lt <= 168 ? 'Achieved' : 'Overdue'; m.ship_wd = getWorkingDays(sh, de); }
            }
            map.set(n.order_number, m);
        } else {
            map.set(n.order_number, n);
        }
    });

    transformedData = Array.from(map.values());
    transformedData.sort((a, b) => b.cohort_date.localeCompare(a.cohort_date));
}

// === Chart Utilities ===
function drawChart(canvasId, type, data, options) {
    if (chartInstances[canvasId]) chartInstances[canvasId].destroy();
    const ctx = document.getElementById(canvasId).getContext('2d');
    chartInstances[canvasId] = new Chart(ctx, { type, data, options });
}

function getDoughnutOptions(centerText, centerSubtext) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '72%',
        plugins: {
            centerText: { text: centerText || '', subtext: centerSubtext || '' },
            tooltip: {
                callbacks: {
                    label: function(ctx) {
                        const total = ctx.dataset.data.reduce((a, v) => a + v, 0);
                        const pct = total > 0 ? Math.round((ctx.parsed / total) * 100) + '%' : '0%';
                        return ' ' + ctx.label + ': ' + pct + ' (' + formatNumber(ctx.parsed) + ')';
                    }
                }
            },
            legend: { position: 'bottom', labels: { usePointStyle: true, padding: 14, font: { size: 11 } } }
        }
    };
}

// === Status Badge Class ===
function getStatusClass(status) {
    if (!status) return 'status-pending';
    const s = status.toLowerCase();
    if (s === 'delivered') return 'status-delivered';
    if (s === 'shipped') return 'status-shipped';
    if (s.includes('cancel')) return 'status-cancelled';
    if (s === 'failed') return 'status-failed';
    return 'status-pending';
}

// ============================================================
// === MAIN DASHBOARD RENDER ===
// ============================================================
function renderDashboard() {
    const data = getFilteredData();
    const active = data.filter(d => !d.status.toLowerCase().includes('cancel'));

    // ---- KPI 1: Total Order ----
    const totalOrder = active.length;
    document.getElementById('kpi-paid').innerText = formatNumber(totalOrder);

    // ---- KPI 4: Ticket Rate ----
    const ticketOrders = active.filter(d => d.has_ticket);
    const ticketRate = active.length > 0 ? ((ticketOrders.length / active.length) * 100).toFixed(1) : '0';
    document.getElementById('kpi-ticket-rate').innerText = ticketRate + '%';

    // ---- KPI 6: Avg Production Time (h) ----
    const shipped = active.filter(d => d.pr_lt != null);
    const avgProd = shipped.length > 0 ? (shipped.reduce((s, d) => s + d.pr_lt, 0) / shipped.length).toFixed(1) : '0';
    document.getElementById('kpi-avg-prod').innerText = avgProd + 'h';



    // ---- KPI: Avg Shipping Time (Working Days) ----
    const deliveredWd = active.filter(d => d.ship_wd != null);
    const avgShipWd = deliveredWd.length > 0 ? (deliveredWd.reduce((s, d) => s + d.ship_wd, 0) / deliveredWd.length).toFixed(1) : '0';
    const elShip = document.getElementById('kpi-avg-ship');
    if (elShip) elShip.innerText = avgShipWd + 'd';
}

// ============================================================
// === UI TOGGLES ===
// ============================================================
const btnThemeToggle = document.getElementById('btn-theme-toggle');
if (btnThemeToggle) {
    btnThemeToggle.addEventListener('click', () => {
        document.body.classList.toggle('light-theme');
        const isLight = document.body.classList.contains('light-theme');
        btnThemeToggle.innerHTML = isLight ? '<i data-lucide="moon"></i>' : '<i data-lucide="sun"></i>';
        lucide.createIcons();
    });
}

const btnSidebarToggle = document.getElementById('btn-sidebar-toggle');
const sidebar = document.querySelector('.sidebar');
if (btnSidebarToggle && sidebar) {
    btnSidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
        const isColl = sidebar.classList.contains('collapsed');
        btnSidebarToggle.innerHTML = isColl ? '<i data-lucide="panel-left-open"></i>' : '<i data-lucide="panel-left-close"></i>';
        lucide.createIcons();
    });
}
