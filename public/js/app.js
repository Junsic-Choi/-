const API_BASE = '/api';
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
    weekPicker.value = currentWeekId;
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
            e.target.value = currentWeekId; // prevent clearing
            return;
        }
        currentWeekId = e.target.value;
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
    if (equipments.length > 0) {
        selectEquipment(equipments[0]); // Select first tab by default
    }
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
        const mJson = await mRes.json();
        if (mJson.success && Array.isArray(mJson.data)) {
            mJson.data.forEach(m => { if (m) allManagersSet.add(m); });
        }

        // 2. Fetch current week for mapping (enables equipment filtering)
        const res = await fetch(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`);
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
    } catch (err) {
        console.error("Data refresh error:", err);
    } finally {
        updateManagerOptions(allManagersSet);
        renderTabs();
    }
}

// Render Tabs
function renderTabs() {
    equipmentTabs.innerHTML = '';
    const selectedManager = managerFilter.value;

    // Equipment Tabs
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
        <td style="width: 4%;"><input type="text" name="mon" value="${data.mon || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="tue" value="${data.tue || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="wed" value="${data.wed || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="thu" value="${data.thu || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="fri" value="${data.fri || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sat" value="${data.sat || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
        <td style="width: 4%;"><input type="text" name="sun" value="${data.sun || ''}" maxlength="1" style="text-align: center; width: 100%;"></td>
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

            const days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

            for (const [eq, plans] of Object.entries(groups)) {
                const equipmentHolidays = holidaysMap[eq] || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };

                // Calculate group rate
                let groupPlanTotal = 0;
                let groupActTotal = 0;

                plans.forEach(p => {
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
                plans.forEach(plan => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><strong>${plan.equipment}</strong></td>
                        <td>${plan.manager}</td>
                        <td>${plan.model}</td>
                        <td>${plan.partName}</td>
                        <td>${plan.partNo}</td>
                        <td class="type-cell">
                            <div class="stats-row center">계획</div>
                            <div class="stats-row center">실적</div>
                        </td>
                        ${getCellHtml(plan, 'mon', equipmentHolidays)}
                        ${getCellHtml(plan, 'tue', equipmentHolidays)}
                        ${getCellHtml(plan, 'wed', equipmentHolidays)}
                        ${getCellHtml(plan, 'thu', equipmentHolidays)}
                        ${getCellHtml(plan, 'fri', equipmentHolidays)}
                        ${getCellHtml(plan, 'sat', equipmentHolidays)}
                        ${getCellHtml(plan, 'sun', equipmentHolidays)}
                    `;
                    consolidatedTableBody.appendChild(tr);
                });
            }
        } else {
            consolidatedTableBody.innerHTML = '<tr><td colspan="13" style="text-align:center;">주간 계획이 등록된 장비가 없습니다.</td></tr>';
        }
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

    let tdClass = '';
    if (isHoliday) tdClass = 'class="holiday-column"';
    else if (isCompleted) tdClass = 'class="completed-cell"';

    return `<td ${tdClass} style="vertical-align: middle;">
        <div class="stats-row"><span class="plan-val-text">${isHoliday ? 'X' : pStr}</span></div>
        <div class="stats-row"><input type="text" class="act-input" data-id="${plan.id}" data-day="${day}_act" value="${aStr}" maxlength="2" ${isHoliday ? 'disabled placeholder="X"' : ''}></div>
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
    // Get current week string like "2026-W08"
    const now = new Date();
    const oneJan = new Date(now.getFullYear(), 0, 1);
    const numberOfDays = Math.floor((now - oneJan) / (24 * 60 * 60 * 1000));
    const weekNumber = Math.ceil((now.getDay() + 1 + numberOfDays) / 7);
    return `${now.getFullYear()}-W${String(weekNumber).padStart(2, '0')}`;
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

    const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    const korDays = ['월', '화', '수', '목', '금', '토', '일'];

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

// 엑셀 추출 기능 (SheetJS 사용) - 기존 양식 참조 고도화
async function exportToExcel() {
    try {
        const res = await fetch(`${API_BASE}/plans-consolidated/${encodeURIComponent(currentWeekId)}`);
        const json = await res.json();

        if (!json.success || json.data.length === 0) {
            showToast('추출할 데이터가 없습니다.', 'error');
            return;
        }

        const hRes = await fetch(`${API_BASE}/holidays-all/${encodeURIComponent(currentWeekId)}`);
        const hJson = await hRes.json();
        const holidaysMap = hJson.data || {};

        // 주차 시작일 계산 (헤더용 날짜 정보)
        const year = parseInt(currentWeekId.substring(0, 4));
        const week = parseInt(currentWeekId.substring(6, 8));
        const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
        const dow = simpleDate.getDay();
        const ISOweekStart = new Date(simpleDate);
        if (dow <= 4)
            ISOweekStart.setDate(simpleDate.getDate() - simpleDate.getDay() + 1);
        else
            ISOweekStart.setDate(simpleDate.getDate() + 8 - simpleDate.getDay());

        const worksheetData = [];

        // 헤더 구성 (기존 양식 참조)
        const headerRow = ['NO', '담당자', '기종', '품명', '품번', '구분', '장비(W/C)'];
        const days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
        const korDays = ['월', '화', '수', '목', '금', '토', '일'];

        days.forEach((d, i) => {
            const targetDate = new Date(ISOweekStart);
            targetDate.setDate(ISOweekStart.getDate() + i);
            const dateStr = `${targetDate.getMonth() + 1}/${targetDate.getDate()}`;
            headerRow.push(`${korDays[i]}(${dateStr})`);
        });
        worksheetData.push(headerRow);

        json.data.forEach((plan, index) => {
            const eq = plan.equipment;
            const h = holidaysMap[eq] || {};

            // 계획 행
            const planRow = [
                index + 1,
                plan.manager || '',
                plan.model || '',
                plan.partName || '',
                plan.partNo || '',
                '계획',
                eq
            ];
            days.forEach(d => {
                planRow.push(h[d] === 1 ? 'X' : (plan[d] || ''));
            });
            worksheetData.push(planRow);

            // 실적 행
            const actRow = [
                '', '', '', '', '', '실적', ''
            ];
            days.forEach(d => {
                actRow.push(h[d] === 1 ? 'X' : (plan[`${d}_act`] || ''));
            });
            worksheetData.push(actRow);
        });

        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

        // 열 너비 조정
        worksheet['!cols'] = [
            { wch: 5 },  // NO
            { wch: 10 }, // 담당자
            { wch: 15 }, // 기종
            { wch: 25 }, // 품명
            { wch: 25 }, // 품번
            { wch: 8 },  // 구분
            { wch: 15 }, // 장비(W/C)
            { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 } // 요일
        ];

        XLSX.utils.book_append_sheet(workbook, worksheet, "통합계획");

        // 다운로드 실행
        XLSX.writeFile(workbook, `정삭계획_통합_${currentWeekId}.xlsx`);
        showToast('기존 양식이 반영된 엑셀 파일이 생성되었습니다.');
    } catch (err) {
        console.error('Export failed:', err);
        alert('엑셀 추출 중 오류가 발생했습니다.');
    }
}

// 엑셀 추출 버튼 이벤트
const exportExcelBtn = document.getElementById('exportExcelBtn');
if (exportExcelBtn) {
    exportExcelBtn.onclick = exportToExcel;
}

// Build Layout
init();
