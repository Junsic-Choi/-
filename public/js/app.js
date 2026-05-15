const API_BASE = '/api';
let equipments = [];
let currentEquipment = null;
let currentWeekId = getInitialWeekId();
let currentHolidays = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
const LOGS_PER_PAGE = 15;
let allLogs = [];
let allActivityLogs = [];

// Phase: Authentication Check
function checkAuth() {
    const token = localStorage.getItem('mps_auth_token');
    if (!token && !window.location.href.includes('login.html')) {
        window.location.href = 'login.html';
    }
    return token;
}

// Initial check
checkAuth();

function logout() {
    localStorage.removeItem('mps_auth_token');
    window.location.href = 'login.html';
}

async function viewLogs() {
    const modal = document.getElementById('logModal');
    const tbody = document.getElementById('logTableBody');
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">로딩 중...</td></tr>';
    modal.classList.remove('hidden');
    
    try {
        const res = await fetchWithAuth(`${API_BASE}/logs`);
        if (!res) return;
        const json = await res.json();
        
        if (json.success) {
            allLogs = json.data;
            renderLogPage(1);
        }
    } catch (err) {
        console.error("Failed to fetch logs", err);
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: red;">로그를 불러오는데 실패했습니다.</td></tr>';
    }
}

function renderLogPage(page) {
    const tbody = document.getElementById('logTableBody');
    const pagination = document.getElementById('logPagination');
    tbody.innerHTML = '';
    
    if (allLogs.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">기록된 로그가 없습니다.</td></tr>';
        pagination.innerHTML = '';
        return;
    }

    const start = (page - 1) * LOGS_PER_PAGE;
    const end = start + LOGS_PER_PAGE;
    const pageLogs = allLogs.slice(start, end);

    pageLogs.forEach(log => {
        const tr = document.createElement('tr');
        const timestamp = isNaN(log.timestamp) ? log.timestamp : Number(log.timestamp);
        const date = new Date(timestamp).toLocaleString('ko-KR');
        tr.innerHTML = `
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; font-size: 0.85rem;">${date}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center;">${log.username}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; font-size: 0.85rem;">${log.ip}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center;">
                <span style="color: ${log.status === 'SUCCESS' ? '#10B981' : '#EF4444'}; font-weight: 600;">${log.status}</span>
            </td>
        `;
        tbody.appendChild(tr);
    });

    renderPaginationButtons(allLogs.length, page, pagination, renderLogPage);
}

function closeLogModal() {
    document.getElementById('logModal').classList.add('hidden');
}

async function viewActivityLogs() {
    const modal = document.getElementById('activityLogModal');
    const tbody = document.getElementById('activityLogTableBody');
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">로딩 중...</td></tr>';
    modal.classList.remove('hidden');
    
    try {
        const res = await fetchWithAuth(`${API_BASE}/activity-logs`);
        if (!res) return;
        const json = await res.json();
        
        if (json.success) {
            allActivityLogs = json.data;
            renderActivityLogPage(1);
        }
    } catch (err) {
        console.error("Failed to fetch activity logs", err);
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; color: red;">로그를 불러오는데 실패했습니다.</td></tr>';
    }
}

function renderActivityLogPage(page) {
    const tbody = document.getElementById('activityLogTableBody');
    const pagination = document.getElementById('activityLogPagination');
    tbody.innerHTML = '';
    
    if (allActivityLogs.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">기록된 활동 로그가 없습니다.</td></tr>';
        pagination.innerHTML = '';
        return;
    }

    const start = (page - 1) * LOGS_PER_PAGE;
    const end = start + LOGS_PER_PAGE;
    const pageLogs = allActivityLogs.slice(start, end);

    pageLogs.forEach(log => {
        const tr = document.createElement('tr');
        const timestamp = isNaN(log.timestamp) ? log.timestamp : Number(log.timestamp);
        const date = new Date(timestamp).toLocaleString('ko-KR');
        tr.innerHTML = `
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; font-size: 0.85rem;">${date}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; font-weight: bold;">${log.username}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; color: #1E3A8A; font-size: 0.85rem;">${log.action}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; font-size: 0.85rem;">${log.details}</td>
            <td style="padding: 0.5rem; border: 1px solid #ddd; text-align: center; font-size: 0.85rem;">${log.ip}</td>
        `;
        tbody.appendChild(tr);
    });

    renderPaginationButtons(allActivityLogs.length, page, pagination, renderActivityLogPage);
}

function renderPaginationButtons(totalCount, currentPage, container, renderFn) {
    container.innerHTML = '';
    const totalPages = Math.ceil(totalCount / LOGS_PER_PAGE);
    if (totalPages <= 1) return;

    const maxButtons = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxButtons / 2));
    let endPage = Math.min(totalPages, startPage + maxButtons - 1);

    if (endPage - startPage + 1 < maxButtons) {
        startPage = Math.max(1, endPage - maxButtons + 1);
    }

    if (startPage > 1) {
        const first = document.createElement('button');
        first.textContent = '«';
        first.onclick = () => renderFn(1);
        container.appendChild(first);
    }

    for (let i = startPage; i <= endPage; i++) {
        const btn = document.createElement('button');
        btn.textContent = i;
        if (i === currentPage) btn.classList.add('active');
        btn.onclick = () => renderFn(i);
        container.appendChild(btn);
    }

    if (endPage < totalPages) {
        const last = document.createElement('button');
        last.textContent = '»';
        last.onclick = () => renderFn(totalPages);
        container.appendChild(last);
    }
}

function closeActivityLogModal() {
    document.getElementById('activityLogModal').classList.add('hidden');
}

async function fetchWithAuth(url, options = {}) {
    const token = localStorage.getItem('mps_auth_token');
    const username = localStorage.getItem('mps_user_name') || 'unknown';
    const headers = {
        ...options.headers,
        'Authorization': `Bearer ${token}`,
        'X-User-Name': encodeURIComponent(username)
    };
    
    const response = await fetch(url, { ...options, headers });
    
    if (response.status === 401) {
        localStorage.removeItem('mps_auth_token');
        window.location.href = 'login.html';
        return null;
    }
    
    return response;
}

// DOM Elements
const equipmentTabs = document.getElementById('equipmentTabs');
const weekPicker = document.getElementById('weekPicker');
const weekDisplay = document.getElementById('weekDisplay');
const managerFilter = document.getElementById('managerFilter');
const btnSave = document.getElementById('saveBtn');
const btnAddRow = document.getElementById('addRowBtn');
const planTableBody = document.getElementById('planTableBody');
const planEditor = document.getElementById('planEditor');
const consolidatedView = document.getElementById('consolidatedView');
const consolidatedTableBody = document.getElementById('consolidatedTableBody');
const currentEquipmentTitle = document.getElementById('currentEquipmentTitle');
const btnRefreshConsolidated = document.getElementById('refreshConsolidatedBtn');
const editModeBtn = document.getElementById('editModeBtn');
const saveConsolidatedBtn = document.getElementById('saveConsolidatedBtn');
const equipmentConsFilter = document.getElementById('equipmentConsFilter');
const toastEl = document.getElementById('toast');

let isConsolidatedEditMode = false;

// Initialize
async function init() {
    weekPicker.value = getDateStringFromWeekId(currentWeekId);
    updateTableHeadersForWeek(currentWeekId);

    // Make the custom text div trigger the native date picker when clicked
    weekDisplay.addEventListener('click', () => {
        try {
            weekPicker.showPicker();
        } catch (e) {
            // Fallback for browsers that don't support showPicker on week input
            weekPicker.focus();
        }
    });

    weekPicker.addEventListener('change', async (e) => {
        if (!e.target.value) {
            e.target.value = getDateStringFromWeekId(currentWeekId); // prevent clearing
            return;
        }
        const newWeekId = getWeekIdFromDate(e.target.value);
        if (newWeekId === currentWeekId) {
            e.target.value = getDateStringFromWeekId(currentWeekId);
            return;
        }
        currentWeekId = newWeekId;
        e.target.value = getDateStringFromWeekId(currentWeekId);
        updateTableHeadersForWeek(currentWeekId);
        await refreshManagerFilterAndTabs(); // Update manager filter and possible tabs for new week
        if (currentEquipment) {
            loadPlans(currentEquipment);
        } else {
            loadConsolidatedPlans();
        }
    });

    managerFilter.addEventListener('change', () => {
        renderTabs(); // Filter tabs based on manager
        applyManagerFilter(); // Filter rows in current editor
    });

    equipmentConsFilter.addEventListener('change', () => {
        loadConsolidatedPlans();
    });

    await fetchEquipments();
    selectConsolidatedView(); // Make Consolidated View the default
}

// Global mapping of which manager is on which equipment this week
let managerEquipmentMap = {};

// Fetch Equipments List
async function fetchEquipments() {
    try {
        const res = await fetchWithAuth(`${API_BASE}/equipments`);
        const json = await res.json();
        if (json.success) {
            equipments = json.data;
            // Populate consolidated equipment filter
            equipmentConsFilter.innerHTML = '<option value="">전체 장비</option>';
            equipments.forEach(eq => {
                const opt = document.createElement('option');
                opt.value = opt.textContent = eq;
                equipmentConsFilter.appendChild(opt);
            });
            await refreshManagerFilterAndTabs();
        }
    } catch (err) {
        console.error("Failed to load equipments", err);
    }
}

// Refresh Manager List and Equipment Tabs based on current week's data
async function refreshManagerFilterAndTabs() {
    console.log("Refreshing manager filter and tabs:", currentWeekId);
    let allManagersSet = new Set();
    managerEquipmentMap = {};

    try {
        // 1. Fetch ALL managers globally to ensure dropdown is never empty
        const mRes = await fetchWithAuth(`${API_BASE}/managers`);
        if (mRes.ok) {
            const mJson = await mRes.json();
            if (mJson.success && Array.isArray(mJson.data)) {
                mJson.data.forEach(m => { if (m) allManagersSet.add(m); });
            }
        } else {
            console.warn("Global managers API failed, status:", mRes.status);
        }

        // 2. Fetch current week for mapping (enables equipment filtering)
        const res = await fetchWithAuth(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`);
        if (res.ok) {
            const json = await res.json();
            if (json.success && Array.isArray(json.data)) {
                json.data.forEach(plan => {
                    if (plan.manager) {
                        allManagersSet.add(plan.manager);
                        if (!managerEquipmentMap[plan.manager]) managerEquipmentMap[plan.manager] = new Set();
                        managerEquipmentMap[plan.manager].add(plan.equipment);
                    }
                });
            }
        } else {
            console.warn("Weekly consolidated API failed, status:", res.status);
        }
    } catch (err) {
        console.error("Data refresh error:", err);
        showToast("데이터 연동 중 오류가 발생했습니다. (담당자 목록 확인 필요)", "error");
    } finally {
        updateManagerOptions(allManagersSet);
        renderTabs();
    }
}

// Render Tabs
function renderTabs() {
    equipmentTabs.innerHTML = '';
    const selectedManager = managerFilter.value;

    // Equipment Tabs - ALWAYS show all to allow additions
    equipments.forEach(eq => {
        const li = document.createElement('li');
        li.textContent = eq;
        li.onclick = () => selectEquipment(eq);
        if (eq === currentEquipment) li.classList.add('active');
        equipmentTabs.appendChild(li);
    });

    // Consolidated Tab
    const consLi = document.createElement('li');
    consLi.textContent = "📊 통합 화면";
    consLi.className = "tab-consolidated";
    consLi.onclick = () => selectConsolidatedView();
    if (!currentEquipment) consLi.classList.add('active');
    equipmentTabs.appendChild(consLi);
}

// Tab Selection
function selectEquipment(equipment) {
    currentEquipment = equipment;
    document.querySelectorAll('.tabs li').forEach(li => {
        li.classList.remove('active');
        if (li.textContent === equipment) li.classList.add('active');
    });

    planEditor.classList.add('active');
    consolidatedView.classList.remove('active');
    currentEquipmentTitle.textContent = `${equipment} 장비 계획 입력`;

    // Do not reset managerFilter.value here, loadPlans will handle applying it.
    loadPlans(equipment);
}

// Phase 18: Holiday Management Logic
async function fetchHolidays(equipment) {
    try {
        const res = await fetchWithAuth(`${API_BASE}/holidays/${encodeURIComponent(equipment)}/${encodeURIComponent(currentWeekId)}`);
        const json = await res.json();
        if (json.success && json.data) {
            currentHolidays = json.data;
        } else {
            currentHolidays = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
        }
        renderHolidayUI();
    } catch (err) {
        console.error("Failed to fetch holidays", err);
    }
}

function renderHolidayUI() {
    const btns = document.querySelectorAll('.holiday-btn');
    btns.forEach(btn => {
        const day = btn.getAttribute('data-day');
        // Force Sunday to be active (holiday) and non-interactive
        if (day === 'sun') {
            currentHolidays[day] = 1;
            btn.classList.add('active');
            btn.style.pointerEvents = 'none';
            btn.style.opacity = '0.7';
            return;
        }
        if (currentHolidays[day] === 1) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

async function toggleHoliday(day) {
    currentHolidays[day] = currentHolidays[day] === 1 ? 0 : 1;
    renderHolidayUI();
    try {
        await fetchWithAuth(`${API_BASE}/holidays`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                equipment: currentEquipment,
                weekId: currentWeekId,
                holidays: currentHolidays
            })
        });
        // After saving, reload current table to apply disabled states
        loadPlans(currentEquipment);
    } catch (err) {
        console.error("Failed to save holiday", err);
        showToast("휴무일 저장에 실패했습니다.");
    }
}


function selectConsolidatedView() {
    currentEquipment = null;
    document.querySelectorAll('.tabs li').forEach(li => {
        li.classList.remove('active');
        if (li.textContent.includes('통합 화면')) li.classList.add('active');
    });

    planEditor.classList.remove('active');
    consolidatedView.classList.add('active');

    loadConsolidatedPlans();
}

// Load Plans for Editor
async function loadPlans(equipment) {
    // managerFilter.value = ''; // Reset filter when switching equipment/week - REMOVED to allow filter persistence
    planTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">로딩중...</td></tr>';

    // Parallelize holiday and plan fetching
    try {
        const [hData, pRes] = await Promise.all([
            fetchHolidays(equipment),
            fetchWithAuth(`${API_BASE}/plans/${encodeURIComponent(equipment)}/${encodeURIComponent(currentWeekId)}`)
        ]);
        
        const json = await pRes.json();
        planTableBody.innerHTML = '';

        if (json.success && json.data.length > 0) {
            json.data.forEach((plan, index) => {
                planTableBody.appendChild(createRow(index + 1, plan));
            });
        } else {
            for (let i = 1; i <= 5; i++) {
                planTableBody.appendChild(createRow(i));
            }
        }
        // Ensure global filter is applied to the newly loaded data
        applyManagerFilter();
        // Phase 18: Apply visual/logic for holidays
        applyHolidayRestrictions();
    } catch (err) {
        console.error(err);
    }
}

// Update Manager Options for Dropdown
function updateManagerOptions(managerSet) {
    console.log("Updating manager options dropdown with Set size:", managerSet.size);
    const currentSelectedManager = managerFilter.value;
    managerFilter.innerHTML = '<option value="">전체 항목</option>';

    // Sort managers alphabetically
    const sortedManagers = Array.from(managerSet).sort();

    sortedManagers.forEach(manager => {
        if (!manager) return;
        const option = document.createElement('option');
        option.value = manager;
        option.textContent = manager;
        managerFilter.appendChild(option);
    });

    if (currentSelectedManager && managerSet.has(currentSelectedManager)) {
        managerFilter.value = currentSelectedManager;
    }
}

// Client Side Manager Filter
function applyManagerFilter() {
    const filterText = managerFilter.value.toLowerCase().trim();
    const rows = planTableBody.querySelectorAll('tr');

    rows.forEach(row => {
        const managerInput = row.querySelector('input[name="manager"]');
        if (!managerInput) return; // Ignore if not a standard row
        const val = managerInput.value.toLowerCase().trim();
        // Show if: 1) filter is empty, or 2) name matches, or 3) name is empty (to allow new entries)
        if (filterText === '' || val.includes(filterText) || val === '') {
            row.style.display = '';
        } else {
            row.style.display = 'none';
        }
    });
}

function applyHolidayRestrictions() {
    const days = ['fri', 'sat', 'sun', 'mon', 'tue', 'wed', 'thu'];

    // Force Sunday to be a holiday as per user request
    currentHolidays['sun'] = 1;

    // Header cells
    days.forEach(day => {
        const thId = 'th' + day.charAt(0).toUpperCase() + day.slice(1);
        const th = document.getElementById(thId);
        if (!th) return;

        const baseLabel = th.dataset.label || th.textContent.split(' ')[0];
        th.dataset.label = baseLabel; // Store for reuse

        if (currentHolidays[day] === 1) {
            th.classList.add('holiday-column');
            th.innerHTML = `${baseLabel}<br><span class="holiday-cell-text">(휴무)</span>`;
        } else {
            th.classList.remove('holiday-column');
            th.innerHTML = baseLabel;
        }
    });

    // Body cells
    const rows = planTableBody.querySelectorAll('tr');
    rows.forEach(row => {
        days.forEach(day => {
            const input = row.querySelector(`input[name="${day}"]`);
            if (!input) return;
            const td = input.parentElement;

            if (currentHolidays[day] === 1) {
                td.classList.add('holiday-column');
                input.disabled = true;
                input.placeholder = "";
                input.value = "";
            } else {
                td.classList.remove('holiday-column');
                input.disabled = false;
                input.placeholder = "";
            }
        });
    });
}

// Create Editable Row
function createRow(index, data = {}) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td style="width: 3%; text-align: center;">${index}</td>
        <td style="width: 8%;"><input type="text" name="manager" value="${data.manager || ''}" placeholder="담당자"></td>
        <td style="width: 12%;"><input type="text" name="model" value="${data.model || ''}" placeholder="기종"></td>
        <td style="width: 22%;"><input type="text" name="partName" value="${data.partName || ''}" placeholder="품명" style="width: 100%;"></td>
        <td style="width: 22%;">
            <input type="text" name="partNo" value="${data.partNo || ''}" placeholder="품번" style="width: 100%;">
            ${data.urgentStatus ? `<div class="urgent-label">${data.urgentStatus}</div>` : ''}
        </td>
        <td style="width: 4%;"><input type="text" name="fri" value="${data.fri || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sat" value="${data.sat || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sun" value="${data.sun || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="mon" value="${data.mon || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="tue" value="${data.tue || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="wed" value="${data.wed || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="thu" value="${data.thu || ''}" maxlength="2" style="text-align: center; width: 100%;"></td>
        <td style="width: 5%;"><button class="btn btn-danger" onclick="this.closest('tr').remove()">삭제</button></td>
    `;
    return tr;
}

// Add New Row Event
btnAddRow.addEventListener('click', () => {
    const nextIndex = planTableBody.children.length + 1;
    planTableBody.appendChild(createRow(nextIndex));
});

// Save Plans Event
btnSave.addEventListener('click', async () => {
    if (!currentEquipment) return;

    const rows = planTableBody.querySelectorAll('tr');
    const plansToSave = [];
    const partNoSet = new Set();

    for (const row of rows) {
        const inputs = row.querySelectorAll('input');
        const plan = {};
        inputs.forEach(input => {
            plan[input.name] = input.value.trim();
        });

        // Only save rows that have at least some basic information or plan data
        const hasData = Object.values(plan).some(v => v !== '');
        if (hasData) {
            const normalizedPartNo = plan.partNo ? plan.partNo.trim().toUpperCase() : '';
            if (normalizedPartNo && partNoSet.has(normalizedPartNo)) {
                alert(`동일 장비 중복 품번입니다: ${plan.partNo}`);
                return;
            }
            if (normalizedPartNo) partNoSet.add(normalizedPartNo);
            plansToSave.push(plan);
        }
    }

    try {
        const res = await fetchWithAuth(`${API_BASE}/plans/${encodeURIComponent(currentEquipment)}/${encodeURIComponent(currentWeekId)}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ plans: plansToSave })
        });
        const json = await res.json();
        if (json.success) {
            showToast(`[${currentEquipment}] 계획이 성공적으로 저장되었습니다! 🎉`);
            loadPlans(currentEquipment); // Reload to format IDs
        } else {
            alert('저장 실패: ' + json.error);
        }
    } catch (err) {
        console.error(err);
        alert('네트워크 오류가 발생했습니다.');
    }
});

// Load Consolidated View
async function loadConsolidatedPlans() {
    consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">데이터를 불러오는 중입니다...</td></tr>';
    try {
        const [res, hRes] = await Promise.all([
            fetchWithAuth(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`),
            fetchWithAuth(`${API_BASE}/holidays-all/${encodeURIComponent(currentWeekId)}`)
        ]);
        
        const json = await res.json();
        const hJson = await hRes.json();
        const holidaysMap = hJson.data || {}; 

        const fragment = document.createDocumentFragment();
        if (json.success && json.data.length > 0) {
            const filterEq = equipmentConsFilter.value;
            const days = ['fri', 'sat', 'sun', 'mon', 'tue', 'wed', 'thu'];

            // Group by equipment
            const groups = {};
            const orderedEquipment = [];
            
            json.data.forEach(plan => {
                const eq = plan.equipment;
                if (filterEq && eq !== filterEq) return;
                
                if (!groups[eq]) {
                    groups[eq] = [];
                    orderedEquipment.push(eq);
                }
                groups[eq].push(plan);
            });

            if (orderedEquipment.length === 0) {
                consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">조건에 맞는 주간 계획이 없습니다.</td></tr>';
                return;
            }

            for (const eq of orderedEquipment) {
                const plans = groups[eq];
                const activePlans = plans.filter(p => days.some(d => p[d] && String(p[d]).trim() !== ''));
                if (activePlans.length === 0) continue; 

                const equipmentHolidays = holidaysMap[eq] || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };

                // Render Group Header
                const headerTr = document.createElement('tr');
                headerTr.className = 'group-header';
                headerTr.innerHTML = `<td colspan="13" style="text-align: left; padding-left: 1.5rem;">${eq}</td>`;
                fragment.appendChild(headerTr);

                // Sort activePlans
                activePlans.sort((a, b) => {
                    const managerA = (a.manager || '').trim();
                    const managerB = (b.manager || '').trim();
                    if (managerA !== managerB) return managerA.localeCompare(managerB, 'ko');
                    const getEarliestIndex = (plan) => {
                        for (let i = 0; i < days.length; i++) {
                            const val = parseInt(plan[days[i]]) || 0;
                            if (val > 0) return i;
                        }
                        return 999;
                    };
                    return getEarliestIndex(a) - getEarliestIndex(b);
                });

                // Render Rows
                activePlans.forEach(plan => {
                    const isUrgent = plan.urgentStatus === 'URGENT';
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><strong>${plan.equipment}</strong></td>
                        <td>${plan.manager}</td>
                        <td>${plan.model}</td>
                        <td>${plan.partName}</td>
                        <td class="part-no-cell" onclick="toggleUrgentStatus('${plan.id}', '${plan.urgentStatus || ''}')">
                            <span class="part-no-text">${plan.partNo}</span>
                        </td>
                        <td class="importance-cell">
                            ${isUrgent ? '<span class="importance-star">*</span>' : ''}
                        </td>
                        ${getCellHtml(plan, 'fri', equipmentHolidays)}
                        ${getCellHtml(plan, 'sat', equipmentHolidays)}
                        ${getCellHtml(plan, 'sun', equipmentHolidays)}
                        ${getCellHtml(plan, 'mon', equipmentHolidays)}
                        ${getCellHtml(plan, 'tue', equipmentHolidays)}
                        ${getCellHtml(plan, 'wed', equipmentHolidays)}
                        ${getCellHtml(plan, 'thu', equipmentHolidays)}
                    `;
                    fragment.appendChild(tr);
                });

                // Add Daily Plan Totals Row
                const dailySums = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
                equipmentHolidays.sun = 1; // Force Sunday as holiday for all
                let totalWeeklyPlan = 0;
                let activeDaysCount = 0;
                days.forEach(d => { if (equipmentHolidays[d] !== 1) activeDaysCount++; });

                activePlans.forEach(p => {
                    days.forEach(d => {
                        if (equipmentHolidays[d] === 1) return; // Ignore holiday data for totals
                        const val = parseInt(p[d]) || 0;
                        dailySums[d] += val;
                        totalWeeklyPlan += val;
                    });
                });

                const averagePerDay = activeDaysCount > 0 ? (totalWeeklyPlan / activeDaysCount) : 0;
                const totalRow = document.createElement('tr');
                totalRow.className = 'group-total-row';

                let totalRowHtml = `<td colspan="6" style="text-align: right; font-weight: bold; background-color: #F8FAFC;">[${eq}] 일별 계획 합계</td>`;
                days.forEach(d => {
                    if (equipmentHolidays[d] === 1) {
                        totalRowHtml += `<td class="grid-cell center holiday-column" style="vertical-align: middle; font-weight: bold; color: #1E3A8A;">
                            <div class="stats-row center" style="font-size: 0.95rem;"></div>
                        </td>`;
                        return;
                    }
                    const sum = dailySums[d];
                    const isOverloaded = sum > averagePerDay && sum > 0;
                    const colorStyle = isOverloaded ? 'color: #EF4444;' : 'color: #1E3A8A;';
                    totalRowHtml += `<td class="grid-cell center" style="background-color: #F8FAFC; vertical-align: middle; font-weight: bold; ${colorStyle}">
                        <div class="stats-row center" style="font-size: 0.95rem;">${sum > 0 ? sum : ''}</div>
                    </td>`;
                });
                totalRow.innerHTML = totalRowHtml;
                fragment.appendChild(totalRow);
            }
            consolidatedTableBody.innerHTML = '';
            consolidatedTableBody.appendChild(fragment);
        } else {
            consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">주간 계획이 등록된 데이터가 없습니다.</td></tr>';
        }

        consolidatedTableBody.closest('table').classList.add('consolidated-table');

    } catch (err) {
        console.error(err);
        consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">데이터를 불러오는 데 실패했습니다.</td></tr>';
    }
}

const getCellHtml = (plan, day, equipmentHolidays) => {
    const isHoliday = equipmentHolidays && equipmentHolidays[day] === 1;
    const pStr = plan[day] || '';
    const aStr = plan[`${day}_act`] || '';
    const pVal = parseInt(pStr) || 0;
    const aVal = parseInt(aStr) || 0;

    const isCompleted = pStr !== '' && aVal >= pVal && pVal > 0;

    let tdClass = 'grid-cell';
    if (isHoliday) tdClass += ' holiday-column';
    else if (isCompleted) tdClass += ' completed-cell';

    if (isConsolidatedEditMode && !isHoliday) {
        return `<td class="${tdClass}" style="vertical-align: middle; padding: 0;">
            <input type="text" class="cons-plan-input" data-id="${plan.id}" data-day="${day}" value="${pStr}" maxlength="2" style="width: 100%; height: 100%; border: none; text-align: center; background: transparent; font-size: 1rem; font-weight: 600; color: var(--primary);">
        </td>`;
    }

    return `<td class="${tdClass}" style="vertical-align: middle; text-align: center;">
        <div class="stats-row plan-row center" style="font-weight: 600; font-size: 0.95rem; color: var(--primary); justify-content: center;">${isHoliday ? '' : pStr}</div>
        <div class="stats-row act-row"><input type="text" class="act-input" data-id="${plan.id}" data-day="${day}_act" value="${aStr}" maxlength="2" ${isHoliday ? 'disabled' : ''}></div>
    </td>`;
};

// Phase 7: Save Actuals
const saveActualsBtn = document.getElementById('saveActualsBtn');
if (saveActualsBtn) {
    saveActualsBtn.addEventListener('click', async () => {
        const inputs = document.querySelectorAll('.act-input');
        const updateMap = {};

        inputs.forEach(input => {
            const id = input.getAttribute('data-id');
            const day = input.getAttribute('data-day');
            const val = input.value.trim();

            if (!updateMap[id]) updateMap[id] = { id: id };
            updateMap[id][day] = val;
        });

        const actualsArray = Object.values(updateMap);

        if (actualsArray.length === 0) {
            showToast('저장할 실적 데이터가 없습니다.');
            return;
        }

        try {
            const res = await fetchWithAuth(`${API_BASE}/plans-actuals`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ actuals: actualsArray })
            });
            const json = await res.json();

            if (json.success) {
                showToast('✅ 실적 데이터가 성공적으로 저장되었습니다!');
                loadConsolidatedPlans(); 
            } else {
                alert('저장 실패: ' + json.error);
            }
        } catch (err) {
            console.error(err);
            alert('네트워크 오류가 발생했습니다.');
        }
    });
}

// Toggle Consolidated Edit Mode
editModeBtn.addEventListener('click', () => {
    isConsolidatedEditMode = !isConsolidatedEditMode;
    if (isConsolidatedEditMode) {
        editModeBtn.textContent = '🔓 수정 종료';
        editModeBtn.classList.remove('btn-warning');
        editModeBtn.classList.add('btn-secondary');
        saveConsolidatedBtn.classList.remove('hidden');
    } else {
        editModeBtn.textContent = '✏️ 수정 모드';
        editModeBtn.classList.remove('btn-secondary');
        editModeBtn.classList.add('btn-warning');
        saveConsolidatedBtn.classList.add('hidden');
        loadConsolidatedPlans(); // Refresh to original view
    }
    loadConsolidatedPlans();
});

// Save updated plans from consolidated view
saveConsolidatedBtn.addEventListener('click', async () => {
    const inputs = document.querySelectorAll('.cons-plan-input');
    const updateMap = {};

    inputs.forEach(input => {
        const id = input.getAttribute('data-id');
        const day = input.getAttribute('data-day');
        const val = input.value.trim();

        if (!updateMap[id]) updateMap[id] = { id: id };
        updateMap[id][day] = val;
    });

    const updatesArray = Object.values(updateMap);

    if (updatesArray.length === 0) {
        showToast('변경할 데이터가 없습니다.', 'error');
        return;
    }

    try {
        const res = await fetchWithAuth(`${API_BASE}/plans-batch-update`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ updates: updatesArray })
        });
        const json = await res.json();

        if (json.success) {
            showToast('✅ 계획 수량이 성공적으로 저장되었습니다!');
            isConsolidatedEditMode = false;
            editModeBtn.textContent = '✏️ 수정 모드';
            editModeBtn.classList.remove('btn-secondary');
            editModeBtn.classList.add('btn-warning');
            saveConsolidatedBtn.classList.add('hidden');
            loadConsolidatedPlans();
        } else {
            alert('저장 실패: ' + json.error);
        }
    } catch (err) {
        console.error(err);
        alert('네트워크 오류가 발생했습니다.');
    }
});

// Refresh Consolidated View
document.getElementById('refreshConsolidatedBtn').addEventListener('click', loadConsolidatedPlans);

// Toggle Urgent Status directly on Part No click
async function toggleUrgentStatus(planId, currentStatus) {
    if (!planId || planId === 'undefined' || planId === 'null') {
        alert('이 항목은 아직 현재 주차에 저장되지 않았습니다.\n계획 또는 실적을 먼저 저장한 후 다시 시도해 주세요.');
        return;
    }

    const nextStatus = currentStatus === 'URGENT' ? '' : 'URGENT';

    try {
        console.log(`Toggling urgent status for ID ${planId} to: ${nextStatus}`);
        const res = await fetchWithAuth(`${API_BASE}/urgent-status/${planId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ urgentStatus: nextStatus })
        });
        
        if (!res.ok) {
            const errorText = await res.text();
            throw new Error(`Server responded with ${res.status}: ${errorText}`);
        }

        const json = await res.json();
        if (json.success) {
            showToast(nextStatus === 'URGENT' ? '✅ 중요 항목으로 지정되었습니다.' : '✅ 중요 항목 지정이 해제되었습니다.');
            loadConsolidatedPlans(); 
        } else {
            alert('업데이트 실패: ' + json.error);
        }
    } catch (err) {
        console.error("Urgent update error:", err);
        alert(`업데이트 중 오류가 발생했습니다.\n상태: ${err.message}`);
    }
}

// Phase: Urgent Status Management
function showUrgentMenu(event, planId, currentStatus) {
    // Deprecated in favor of toggleUrgentStatus
}

async function updateUrgentStatus(planId, status) {
    if (!planId || planId === 'undefined' || planId === 'null') {
        alert('이 항목은 아직 현재 주차에 저장되지 않았습니다.\n계획 또는 실적을 먼저 저장한 후 다시 시도해 주세요.');
        return;
    }

    try {
        console.log(`Sending urgent update for ID ${planId} with status: ${status}`);
        const res = await fetchWithAuth(`${API_BASE}/urgent-status/${planId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ urgentStatus: status })
        });
        
        if (!res.ok) {
            const errorText = await res.text();
            throw new Error(`Server responded with ${res.status}: ${errorText}`);
        }

        const json = await res.json();
        if (json.success) {
            showToast('✅ 우선순위 정보가 업데이트되었습니다.');
            loadConsolidatedPlans(); 
        } else {
            alert('업데이트 실패: ' + json.error);
        }
    } catch (err) {
        console.error("Urgent update error:", err);
        alert(`업데이트 중 오류가 발생했습니다.\n상태: ${err.message}`);
    }
}

// Helpers
function highlightPlan(value) {
    if (!value) return '-';
    return `<span style="color: var(--primary); font-weight: 600;">${value}</span>`;
}

function showToast(message) {
    toastEl.textContent = message;
    toastEl.classList.remove('hidden');
    setTimeout(() => {
        toastEl.classList.add('hidden');
    }, 3000);
}

function getInitialWeekId() {
    return getWeekIdFromDate(new Date());
}

function getDateStringFromWeekId(weekString) {
    if (!weekString) return "";
    const year = parseInt(weekString.substring(0, 4));
    const week = parseInt(weekString.substring(6, 8));
    if (isNaN(year) || isNaN(week)) return "";

    const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
    const dow = simpleDate.getDay();
    const ISOweekStart = simpleDate;
    if (dow <= 4)
        ISOweekStart.setDate(simpleDate.getDate() - simpleDate.getDay() + 1);
    else
        ISOweekStart.setDate(simpleDate.getDate() + 8 - simpleDate.getDay());

    ISOweekStart.setDate(ISOweekStart.getDate() - 3); // Friday

    const y = ISOweekStart.getFullYear();
    const m = String(ISOweekStart.getMonth() + 1).padStart(2, '0');
    const d = String(ISOweekStart.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function getWeekIdFromDate(dateInput) {
    let selectedDate;
    if (typeof dateInput === 'string' && dateInput.includes('-')) {
        const parts = dateInput.split('-');
        selectedDate = new Date(parts[0], parts[1] - 1, parts[2]);
    } else {
        selectedDate = new Date(dateInput);
    }
    if (isNaN(selectedDate)) return null;

    const day = selectedDate.getDay();
    // Week starts on Friday. Offset to go back to the most recent Friday.
    // Fri(5)->0, Sat(6)->1, Sun(0)->2, Mon(1)->3, Tue(2)->4, Wed(3)->5, Thu(4)->6
    const offset = (day + 2) % 7;
    selectedDate.setDate(selectedDate.getDate() - offset);

    // To get the ISO week ID for a Friday-Thursday week,
    // we use the Thursday (the last day of the week) for the calculation.
    const thursday = new Date(selectedDate);
    thursday.setDate(thursday.getDate() + 6);

    const d = new Date(Date.UTC(thursday.getFullYear(), thursday.getMonth(), thursday.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function updateTableHeadersForWeek(weekString) {
    // weekString format: "2026-W08"
    if (!weekString) return;

    // Convert ISO week to start date (Monday)
    const year = parseInt(weekString.substring(0, 4));
    const week = parseInt(weekString.substring(6, 8));

    // Simple calculation for week start date (ISO standard)
    const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
    const dow = simpleDate.getDay();
    const ISOweekStart = simpleDate;
    if (dow <= 4)
        ISOweekStart.setDate(simpleDate.getDate() - simpleDate.getDay() + 1);
    else
        ISOweekStart.setDate(simpleDate.getDate() + 8 - simpleDate.getDay());

    // Start from Friday
    ISOweekStart.setDate(ISOweekStart.getDate() - 3);

    const days = ['Fri', 'Sat', 'Sun', 'Mon', 'Tue', 'Wed', 'Thu'];
    const korDays = ['금', '토', '일', '월', '화', '수', '목'];

    let startDateStr = "";
    let endDateStr = "";

    for (let i = 0; i < 7; i++) {
        const targetDate = new Date(ISOweekStart);
        targetDate.setDate(ISOweekStart.getDate() + i);
        const formattedDate = `${targetDate.getMonth() + 1}/${targetDate.getDate()}`;

        if (i === 0) startDateStr = formattedDate;
        if (i === 6) endDateStr = formattedDate;

        const thEq = document.getElementById(`th${days[i]}`);
        const thCon = document.getElementById(`thCons${days[i]}`);

        if (thEq) thEq.textContent = `${korDays[i]} (${formattedDate})`;
        if (thCon) thCon.textContent = `${korDays[i]} (${formattedDate})`;
    }

    if (weekDisplay) {
        weekDisplay.textContent = `📅 ${startDateStr} ~ ${endDateStr}`;
    }
}


// Phase 18: Holiday Toggle Click
const holidayToggles = document.getElementById('holidayToggles');
if (holidayToggles) {
    holidayToggles.onclick = (e) => {
        if (e.target.classList.contains('holiday-btn')) {
            const day = e.target.getAttribute('data-day');
            toggleHoliday(day);
        }
    };
}

// 엑셀 추출 기능 (SheetJS 사용) - 서버리스 환경 호환성 및 인증 헤더 대응
async function exportToExcel() {
    try {
        showToast('엑셀 추출을 준비 중입니다...');
        
        // 1. 서버 사이드 스타일 엑셀 시도 (인증 헤더 포함)
        const url = `${API_BASE}/export-excel-styled/${encodeURIComponent(currentWeekId)}`;
        const response = await fetchWithAuth(url);
        
        // 2. 만약 서버에서 지원하지 않거나(501) 오류가 나면 프론트엔드 SheetJS로 백업
        if (response && response.ok) {
            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = `Integrated_Plan_${currentWeekId}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(downloadUrl);
            a.remove();
            showToast('엑셀 파일이 다운로드 되었습니다.');
        } else {
            console.warn("Server-side export failed or not supported, falling back to SheetJS");
            // SheetJS Backup
            const table = document.getElementById('consolidatedTable');
            if (!table) throw new Error("데이터 테이블을 찾을 수 없습니다.");
            
            const wb = XLSX.utils.table_to_book(table, { sheet: "통합계획" });
            XLSX.writeFile(wb, `Integrated_Plan_${currentWeekId}.xlsx`);
            showToast('엑셀 파일이 다운로드 되었습니다. (기본 양식)');
        }
    } catch (err) {
        console.error('Export failed:', err);
        alert(`엑셀 추출 중 오류가 발생했습니다: ${err.message}`);
    }
}

// 엑셀 추출 버튼 이벤트
const exportExcelBtn = document.getElementById('exportExcelBtn');
if (exportExcelBtn) {
    exportExcelBtn.onclick = exportToExcel;
}

// Build Layout
init();
