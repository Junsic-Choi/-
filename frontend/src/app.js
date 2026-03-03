

const API_BASE = import.meta.env.VITE_API_URL || '/api';
let equipments = [];
let currentEquipment = null;
let currentWeekId = getInitialWeekId();
let currentHolidays = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };

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
const toastEl = document.getElementById('toast');

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

    await fetchEquipments();
    selectConsolidatedView(); // Make Consolidated View the default
}

// Global mapping of which manager is on which equipment this week
let managerEquipmentMap = {};

// Fetch Equipments List
async function fetchEquipments() {
    try {
        const res = await fetch(`${API_BASE}/equipments`);
        const json = await res.json();
        if (json.success) {
            equipments = json.data;
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
        const mRes = await fetch(`${API_BASE}/managers`);
        if (mRes.ok) {
            const mJson = await mRes.json();
            if (mJson.success && Array.isArray(mJson.data)) {
                mJson.data.forEach(m => { if (m) allManagersSet.add(m); });
            }
        } else {
            console.warn("Global managers API failed, status:", mRes.status);
        }

        // 2. Fetch current week for mapping (enables equipment filtering)
        const res = await fetch(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`);
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
        const res = await fetch(`${API_BASE}/holidays/${encodeURIComponent(equipment)}/${encodeURIComponent(currentWeekId)}`);
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
        await fetch(`${API_BASE}/holidays`, {
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

    // Fetch holidays first
    await fetchHolidays(equipment);

    try {
        const res = await fetch(`${API_BASE}/plans/${encodeURIComponent(equipment)}/${encodeURIComponent(currentWeekId)}`);
        const json = await res.json();
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
    const days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

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
                input.placeholder = "X";
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
        <td style="width: 22%;"><input type="text" name="partNo" value="${data.partNo || ''}" placeholder="품번" style="width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sun" value="${data.sun || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="mon" value="${data.mon || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="tue" value="${data.tue || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="wed" value="${data.wed || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="thu" value="${data.thu || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="fri" value="${data.fri || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sat" value="${data.sat || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
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

    rows.forEach(row => {
        const inputs = row.querySelectorAll('input');
        const plan = {};
        inputs.forEach(input => {
            plan[input.name] = input.value.trim();
        });

        // Only save rows that have at least some basic information or plan data
        const hasData = Object.values(plan).some(v => v !== '');
        if (hasData) {
            plansToSave.push(plan);
        }
    });

    try {
        const res = await fetch(`${API_BASE}/plans/${encodeURIComponent(currentEquipment)}/${encodeURIComponent(currentWeekId)}`, {
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
        const res = await fetch(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`);
        const json = await res.json();

        // Phase 18: Fetch all holidays for this week
        const hRes = await fetch(`${API_BASE}/holidays-all/${encodeURIComponent(currentWeekId)}`);
        const hJson = await hRes.json();
        const holidaysMap = hJson.data || {}; // { equipment: { mon: 1, ... } }

        consolidatedTableBody.innerHTML = '';
        if (json.success && json.data.length > 0) {
            // Group by equipment
            const groups = {};
            json.data.forEach(plan => {
                const eq = plan.equipment;
                if (!groups[eq]) groups[eq] = [];
                groups[eq].push(plan);
            });

            const days = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

            for (const [eq, plans] of Object.entries(groups)) {
                // Filter plans to only those with data (Actual planning hours must be present)
                const activePlans = plans.filter(p => ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'].some(d => p[d] && String(p[d]).trim() !== ''));
                if (activePlans.length === 0) continue; // Skip equipment group if no active plans

                const equipmentHolidays = holidaysMap[eq] || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };

                // Calculate group rate
                let groupPlanTotal = 0;
                let groupActTotal = 0;

                activePlans.forEach(p => {
                    days.forEach(d => {
                        groupPlanTotal += parseInt(p[d]) || 0;
                        groupActTotal += parseInt(p[`${d}_act`]) || 0;
                    });
                });

                const rate = groupPlanTotal > 0 ? Math.round((groupActTotal / groupPlanTotal) * 100) : 0;

                // Render Group Header
                const headerTr = document.createElement('tr');
                headerTr.className = 'group-header';
                headerTr.innerHTML = `<td colspan="13">${eq} (총 실적률: ${rate}%)</td>`;
                consolidatedTableBody.appendChild(headerTr);

                // Render Rows
                activePlans.forEach(plan => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><strong>${plan.equipment}</strong></td>
                        <td>${plan.manager}</td>
                        <td>${plan.model}</td>
                        <td>${plan.partName}</td>
                        <td>${plan.partNo}</td>
                        <td class="type-cell grid-cell">
                            <div class="stats-row plan-row center">계획</div>
                            <div class="stats-row act-row center">실적</div>
                        </td>
                        ${getCellHtml(plan, 'sun', equipmentHolidays)}
                        ${getCellHtml(plan, 'mon', equipmentHolidays)}
                        ${getCellHtml(plan, 'tue', equipmentHolidays)}
                        ${getCellHtml(plan, 'wed', equipmentHolidays)}
                        ${getCellHtml(plan, 'thu', equipmentHolidays)}
                        ${getCellHtml(plan, 'fri', equipmentHolidays)}
                        ${getCellHtml(plan, 'sat', equipmentHolidays)}
                    `;
                    consolidatedTableBody.appendChild(tr);
                });

                // Phase: Add Daily Plan Totals Row (Excluding Actuals)
                const totalRow = document.createElement('tr');
                totalRow.className = 'group-total-row';

                // Calculate daily plan sums
                const dailySums = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
                activePlans.forEach(p => {
                    days.forEach(d => {
                        dailySums[d] += parseInt(p[d]) || 0;
                    });
                });

                let totalRowHtml = `
                    <td colspan="5" style="text-align: right; font-weight: bold; background-color: #F8FAFC;">[${eq}] 일별 계획 합계</td>
                    <td class="type-cell grid-cell" style="background-color: #F8FAFC;">
                        <div class="stats-row center" style="font-weight: bold;">합계</div>
                    </td>
                `;

                days.forEach(d => {
                    totalRowHtml += `<td class="grid-cell center" style="background-color: #F8FAFC; vertical-align: middle; font-weight: bold; color: #1E3A8A;">
                        <div class="stats-row center">${dailySums[d] > 0 ? dailySums[d] : ''}</div>
                    </td>`;
                });

                totalRow.innerHTML = totalRowHtml;
                consolidatedTableBody.appendChild(totalRow);
            }
        } else {
            consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">주간 계획이 등록된 장비가 없습니다.</td></tr>';
        }


        // Add grid class to table
        consolidatedTableBody.closest('table').classList.add('consolidated-table');

        // Apply yellow highlight to completed cells
        document.querySelectorAll('.consolidated-table .completed-cell').forEach(td => {
            td.style.backgroundColor = '#FFFF99'; // Excel-like Yellow
        });

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

    return `<td class="${tdClass}" style="vertical-align: middle;">
        <div class="stats-row plan-row"><span class="plan-val-text">${isHoliday ? 'X' : pStr}</span></div>
        <div class="stats-row act-row"><input type="text" class="act-input" data-id="${plan.id}" data-day="${day}_act" value="${aStr}" maxlength="2" ${isHoliday ? 'disabled placeholder="X"' : ''}></div>
    </td>`;
};

// Refresh Consolidated View
document.getElementById('refreshConsolidatedBtn').addEventListener('click', loadConsolidatedPlans);

// Phase 7: Save Actuals
document.getElementById('saveActualsBtn').addEventListener('click', async () => {
    const inputs = document.querySelectorAll('.act-input');
    const updateMap = {};

    inputs.forEach(input => {
        const id = input.getAttribute('data-id');
        const day = input.getAttribute('data-day'); // e.g. 'mon_act'
        const val = input.value.trim();

        if (!updateMap[id]) updateMap[id] = { id: id };
        updateMap[id][day] = val;
    });

    const actualsArray = Object.values(updateMap);

    if (actualsArray.length === 0) {
        showToast('저장할 실적 데이터가 없습니다.', 'error');
        return;
    }

    try {
        const res = await fetch(`${API_BASE}/plans-actuals`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ actuals: actualsArray })
        });
        const json = await res.json();

        if (json.success) {
            showToast('✅ 실적 데이터가 성공적으로 저장되었습니다!');
            loadConsolidatedPlans(); // refresh to recalculate rates/colors
        } else {
            alert('저장 실패: ' + json.error);
        }
    } catch (err) {
        console.error(err);
        alert('네트워크 오류가 발생했습니다.');
    }
});

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

    ISOweekStart.setDate(ISOweekStart.getDate() - 1); // Sunday

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
    selectedDate.setDate(selectedDate.getDate() - day);

    const isoMonday = new Date(selectedDate);
    isoMonday.setDate(isoMonday.getDate() + 1);

    const d = new Date(Date.UTC(isoMonday.getFullYear(), isoMonday.getMonth(), isoMonday.getDate()));
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

    // ISO week calculation always gives us a Monday.
    // To make our week start on Sunday, we subtract 1 day from the ISO Monday.
    ISOweekStart.setDate(ISOweekStart.getDate() - 1);

    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const korDays = ['일', '월', '화', '수', '목', '금', '토'];

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

// 엑셀 추출 기능 (SheetJS 사용) - Cloudflare Edge 환경 지원을 위한 프론트엔드 처리 복구
async function exportToExcel() {
    try {
        const wb = XLSX.utils.table_to_book(document.getElementById('consolidatedTable'), { sheet: "통합계획" });
        XLSX.writeFile(wb, `Integrated_Plan_${currentWeekId}.xlsx`);
        showToast('엑셀 파일이 다운로드 되었습니다.');
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

// Initialize Layout immediately without checking auth (Cloudflare Access handles it)
init();
