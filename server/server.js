const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// Initialize SQLite Database
const dbPath = path.resolve(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error('Error opening database', err);
    } else {
        console.log('Database connected.');
        // Create table
        db.run(`CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment TEXT NOT NULL,
            weekId TEXT NOT NULL DEFAULT '2026-W08',
            manager TEXT,
            model TEXT,
            partName TEXT,
            partNo TEXT,
            mon TEXT,
            tue TEXT,
            wed TEXT,
            thu TEXT,
            fri TEXT,
            sat TEXT,
            sun TEXT,
            updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
        )`, (err) => {
            if (err) console.error(err);
        });

        // Ensure weekId column exists if table was already created in Phase 1
        db.run(`ALTER TABLE plans ADD COLUMN weekId TEXT NOT NULL DEFAULT '2026-W08'`, (err) => { /* Ignore */ });

        // Phase 7: Add Actuals columns
        ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'].forEach(day => {
            db.run(`ALTER TABLE plans ADD COLUMN ${day}_act TEXT DEFAULT ''`, (err) => { /* Ignore duplicate column errs */ });
        });

        // Phase 18: Add equipment_holidays table
        db.run(`CREATE TABLE IF NOT EXISTS equipment_holidays (
            equipment TEXT NOT NULL,
            weekId TEXT NOT NULL,
            mon INTEGER DEFAULT 0,
            tue INTEGER DEFAULT 0,
            wed INTEGER DEFAULT 0,
            thu INTEGER DEFAULT 0,
            fri INTEGER DEFAULT 0,
            sat INTEGER DEFAULT 0,
            sun INTEGER DEFAULT 0,
            PRIMARY KEY (equipment, weekId)
        )`, (err) => {
            if (err) console.error(err);
        });

        // Phase 6: Update old equipment names
        db.run(`UPDATE plans SET equipment = 'HSP8000 #1' WHERE equipment = 'HSP8000'`, (err) => {
            if (err) console.error(err);
        });
        db.run(`UPDATE plans SET equipment = 'HSP8000 #2' WHERE equipment = '#2'`, (err) => {
            if (err) console.error(err);
        });
    }
});

// API Routes

// 1. Get equipments list (Hardcoded or distinct from DB - let's offer a fixed list for UI)
// Based on Excel "Holiday" sheet: HSP6300, HSP8000, HM2J, AH2J, Y10T, Y15T, YBM1530
const ALL_EQUIPMENTS = [
    "HSP6300", "HSP8000 #1", "HSP8000 #2", "HM2J", "AH2J", "Y10T", "Y15T", "YBM1530"
];

app.get('/api/equipments', (req, res) => {
    res.json({ success: true, data: ALL_EQUIPMENTS });
});

// 2. Get all distinct managers from the database
app.get('/api/managers', (req, res) => {
    console.log(`[GET] /api/managers requested`);
    db.all(`SELECT DISTINCT manager FROM plans WHERE manager IS NOT NULL AND manager != '' ORDER BY manager ASC`, [], (err, rows) => {
        if (err) {
            console.error('Database error [managers]:', err.message);
            return res.status(500).json({ success: false, error: err.message });
        }
        const managers = rows.map(row => row.manager);
        console.log(`[GET] /api/managers returned ${managers.length} managers`);
        res.json({ success: true, data: managers });
    });
});

// 3. Get plans for a specific equipment AND weekId
app.get('/api/plans/:equipment/:weekId', (req, res) => {
    const { equipment, weekId } = req.params;

    // First, try to get plans for the current requested week
    db.all(`SELECT * FROM plans WHERE equipment = ? AND weekId = ? ORDER BY id ASC`, [equipment, weekId], (err, rows) => {
        if (err) {
            return res.status(500).json({ success: false, error: err.message });
        }

        // If we found data, return it
        if (rows && rows.length > 0) {
            return res.json({ success: true, data: rows });
        }

        // If no data for this week, fetch the most recent data for this equipment
        // Sort by weekId descending (e.g., '2026-W08' > '2026-W07') to get the latest past week
        // ADDED: weekId <= ? to prevent future data from carrying over
        db.all(`SELECT * FROM plans WHERE equipment = ? AND weekId <= ? ORDER BY weekId DESC LIMIT 20`, [equipment, weekId], (err, pastRows) => {
            if (err) {
                return res.status(500).json({ success: false, error: err.message });
            }

            if (pastRows && pastRows.length > 0) {
                // Determine the most recent weekId from the past rows
                const mostRecentWeekId = pastRows[0].weekId;
                // Filter rows just for that most recent week
                const latestWeekRows = pastRows.filter(r => r.weekId === mostRecentWeekId);

                // Create placeholder rows, wiping out the day-specific values
                const carryoverData = latestWeekRows.map(row => ({
                    ...row,
                    id: undefined, // Let the frontend/save layer handle new ids
                    weekId: weekId, // Assign to the newly requested week
                    mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "",
                    mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: ""
                }));
                return res.json({ success: true, data: carryoverData });
            }

            // Absolutely no data found previously or currently
            res.json({ success: true, data: [] });
        });
    });
});

// 3. Save plans for a specific equipment AND weekId
// This overrides existing plans for that equipment/weekId and sets new ones.
// It ALSO synchronizes these plans to all existing future weeks for this equipment.
app.post('/api/plans/:equipment/:weekId', (req, res) => {
    const { equipment, weekId } = req.params;
    const plans = req.body.plans; // Array of plan objects

    db.serialize(() => {
        db.run(`BEGIN TRANSACTION;`);

        // --- 1. Save plans for the CURRENT week ---
        db.run(`DELETE FROM plans WHERE equipment = ? AND weekId = ?`, [equipment, weekId], (err) => {
            if (err) {
                console.error("Delete Error", err);
                db.run(`ROLLBACK;`);
                return res.status(500).json({ success: false, error: err.message });
            }
        });

        const stmt = db.prepare(`INSERT INTO plans 
            (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

        plans.forEach(p => {
            stmt.run([
                equipment, weekId, p.manager || "", p.model || "", p.partName || "", p.partNo || "",
                p.mon || "", p.tue || "", p.wed || "", p.thu || "", p.fri || "", p.sat || "", p.sun || ""
            ]);
        });
        stmt.finalize();

        // --- 2. Synchronize plans to FUTURE weeks ---
        // Find all unique future weeks for this equipment
        db.all(`SELECT DISTINCT weekId FROM plans WHERE equipment = ? AND weekId > ? ORDER BY weekId ASC`, [equipment, weekId], (err, futureWeeksRows) => {
            if (err) {
                console.error("Future Weeks Fetch Error", err);
                db.run(`ROLLBACK;`);
                return res.status(500).json({ success: false, error: err.message });
            }

            if (futureWeeksRows.length === 0) {
                // No future weeks to sync, just commit
                db.run(`COMMIT;`, (err) => {
                    if (err) return res.status(500).json({ success: false, error: err.message });
                    return res.json({ success: true, message: 'Plans saved successfully.' });
                });
                return;
            }

            // Sync each future week
            let completedWeeks = 0;
            const totalWeeks = futureWeeksRows.length;

            futureWeeksRows.forEach(row => {
                const targetWeek = row.weekId;

                // Fetch existing data for the target week to preserve plan/actual values
                db.all(`SELECT * FROM plans WHERE equipment = ? AND weekId = ?`, [equipment, targetWeek], (err, targetRows) => {
                    if (err) {
                        db.run(`ROLLBACK;`);
                        return res.status(500).json({ success: false, error: err.message });
                    }

                    // Delete existing rows for target week
                    db.run(`DELETE FROM plans WHERE equipment = ? AND weekId = ?`, [equipment, targetWeek], (err) => {
                        if (err) {
                            db.run(`ROLLBACK;`);
                            return res.status(500).json({ success: false, error: err.message });
                        }

                        // Insert the newly synced item list
                        const syncStmt = db.prepare(`INSERT INTO plans 
                            (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun, mon_act, tue_act, wed_act, thu_act, fri_act, sat_act, sun_act) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

                        plans.forEach(p => {
                            // Try to find if this exact item existed in the target week
                            const existingRow = targetRows.find(tr =>
                                tr.partNo === p.partNo && tr.partName === p.partName && tr.model === p.model
                            );

                            syncStmt.run([
                                equipment, targetWeek, p.manager || "", p.model || "", p.partName || "", p.partNo || "",
                                existingRow ? existingRow.mon : "",
                                existingRow ? existingRow.tue : "",
                                existingRow ? existingRow.wed : "",
                                existingRow ? existingRow.thu : "",
                                existingRow ? existingRow.fri : "",
                                existingRow ? existingRow.sat : "",
                                existingRow ? existingRow.sun : "",
                                existingRow ? existingRow.mon_act : "",
                                existingRow ? existingRow.tue_act : "",
                                existingRow ? existingRow.wed_act : "",
                                existingRow ? existingRow.thu_act : "",
                                existingRow ? existingRow.fri_act : "",
                                existingRow ? existingRow.sat_act : "",
                                existingRow ? existingRow.sun_act : ""
                            ]);
                        });
                        syncStmt.finalize();

                        completedWeeks++;
                        if (completedWeeks === totalWeeks) {
                            // All future weeks synced, commit
                            db.run(`COMMIT;`, (err) => {
                                if (err) return res.status(500).json({ success: false, error: err.message });
                                return res.json({ success: true, message: 'Plans and future weeks synchronized successfully.' });
                            });
                        }
                    });
                });
            });
        });
    });
});

// Phase 18: Holidays API
app.get('/api/holidays/:equipment/:weekId', (req, res) => {
    const { equipment, weekId } = req.params;
    db.get(`SELECT * FROM equipment_holidays WHERE equipment = ? AND weekId = ?`, [equipment, weekId], (err, row) => {
        if (err) return res.status(500).json({ success: false, error: err.message });
        res.json({ success: true, data: row || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 } });
    });
});

// 4. Get consolidated plans per weekId (with Robust Per-Equipment Carryover)
app.get('/api/plans-consolidated/:weekId', (req, res) => {
    const { weekId } = req.params;
    console.log(`[GET] /api/plans-consolidated/${weekId} requested`);

    // Fetch ONLY the latest week's data per equipment, rather than ALL historical rows
    const sql = `
        SELECT p.* 
        FROM plans p
        INNER JOIN (
            SELECT equipment, MAX(weekId) as maxWeek
            FROM plans
            WHERE weekId <= ?
            GROUP BY equipment
        ) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek
    `;

    db.all(sql, [weekId], (err, rows) => {
        if (err) {
            console.error('Database error [plans-consolidated]:', err.message);
            return res.status(500).json({ success: false, error: err.message });
        }

        const consolidatedData = [];

        // Process the rows to clear carryover data fields
        const processedData = rows.map(row => {
            if (row.weekId === weekId) {
                return row;
            } else {
                return {
                    ...row,
                    id: undefined,
                    weekId: weekId,
                    mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "",
                    mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: ""
                };
            }
        });

        // Re-apply ALL_EQUIPMENTS ordering so the UI doesn't scramble the rows
        ALL_EQUIPMENTS.forEach(eq => {
            const machineData = processedData.filter(d => d.equipment.trim() === eq.trim());
            machineData.forEach(row => consolidatedData.push(row));
        });

        console.log(`[GET] /api/plans-consolidated/${weekId} returned ${consolidatedData.length} rows (Optimized Mixed current/carryover)`);
        res.json({ success: true, data: consolidatedData });
    });
});

app.get('/api/holidays-all/:weekId', (req, res) => {
    const { weekId } = req.params;
    console.log(`[GET] /api/holidays-all/${weekId} requested`);
    db.all(`SELECT * FROM equipment_holidays WHERE weekId = ?`, [weekId], (err, rows) => {
        if (err) {
            console.error('Database error [holidays-all]:', err.message);
            return res.status(500).json({ success: false, error: err.message });
        }
        const map = {};
        rows.forEach(r => {
            map[r.equipment] = r;
        });
        console.log(`[GET] /api/holidays-all/${weekId} returned ${rows.length} holiday settings`);
        res.json({ success: true, data: map });
    });
});

app.post('/api/holidays', (req, res) => {
    const { equipment, weekId, holidays } = req.body;
    const { mon, tue, wed, thu, fri, sat, sun } = holidays;
    const sql = `
        INSERT INTO equipment_holidays (equipment, weekId, mon, tue, wed, thu, fri, sat, sun)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(equipment, weekId) DO UPDATE SET
            mon=excluded.mon, tue=excluded.tue, wed=excluded.wed, 
            thu=excluded.thu, fri=excluded.fri, sat=excluded.sat, sun=excluded.sun
    `;
    db.run(sql, [equipment, weekId, mon, tue, wed, thu, fri, sat, sun], function (err) {
        if (err) return res.status(500).json({ success: false, error: err.message });
        res.json({ success: true, message: 'Holidays updated successfully.' });
    });
});

// 5. Save actuals from consolidated view
// Expects body: { actuals: [{id: 1, mon_act: '5', tue_act: '3', ...}, ...] }
// Phase 29: Stylish Excel Export
const ExcelJS = require('exceljs');

function getTextWidth(text) {
    if (text === null || text === undefined) return 0;
    const str = String(text);
    let width = 0;
    for (let i = 0; i < str.length; i++) {
        const charCode = str.charCodeAt(i);
        if (charCode > 255) width += 2; // Korean/Unicode
        else width += 1;
    }
    return width;
}

app.get('/api/export-excel-styled/:weekId', async (req, res) => {
    const { weekId } = req.params;
    console.log(`[GET] /api/export-excel-styled/${weekId} requested`);

    try {
        // Fetch data (reuse optimized SQL logic)
        const sql = `
            SELECT p.* 
            FROM plans p
            INNER JOIN (
                SELECT equipment, MAX(weekId) as maxWeek
                FROM plans
                WHERE weekId <= ?
                GROUP BY equipment
            ) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek
        `;

        db.all(sql, [weekId], async (err, rows) => {
            if (err) return res.status(500).json({ success: false, error: err.message });

            db.all(`SELECT * FROM equipment_holidays WHERE weekId = ?`, [weekId], async (err, holidayRows) => {
                if (err) return res.status(500).json({ success: false, error: err.message });

                const holidaysMap = {};
                holidayRows.forEach(r => holidaysMap[r.equipment] = r);

                const data = [];
                // Process the rows to clear carryover data fields
                const processedData = rows.map(row => {
                    if (row.weekId === weekId) {
                        return row;
                    } else {
                        return {
                            ...row, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: ""
                        };
                    }
                });

                // Re-apply ALL_EQUIPMENTS ordering
                ALL_EQUIPMENTS.forEach(eq => {
                    const machineData = processedData.filter(d => d.equipment.trim() === eq.trim());
                    machineData.forEach(row => data.push(row));
                });

                if (data.length === 0) {
                    return res.status(404).json({ success: false, message: 'No data to export' });
                }

                // Create Workbook
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('통합계획');

                // Page Setup for Print
                worksheet.pageSetup.fitToPage = true;
                worksheet.pageSetup.fitToWidth = 1;
                worksheet.pageSetup.fitToHeight = 0; // 0 means it will scroll pages vertically as needed
                worksheet.pageSetup.orientation = 'landscape';
                worksheet.pageSetup.margins = {
                    left: 0.25, right: 0.25,
                    top: 0.75, bottom: 0.75,
                    header: 0.3, footer: 0.3
                };

                // Header Row Configuration
                const year = parseInt(weekId.substring(0, 4));
                const week = parseInt(weekId.substring(6, 8));

                // Calculate ISO Monday date
                const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
                const dow = simpleDate.getDay();
                const ISOweekStart = new Date(simpleDate);
                if (dow <= 4) ISOweekStart.setDate(simpleDate.getDate() - simpleDate.getDay() + 1);
                else ISOweekStart.setDate(simpleDate.getDate() + 8 - simpleDate.getDay());

                // Shift back 1 day to make Sunday the start of our customized week
                ISOweekStart.setDate(ISOweekStart.getDate() - 1);

                const korDays = ['일', '월', '화', '수', '목', '금', '토'];
                const headerTitles = ['NO', '담당자', '기종', '품명', '품번', '구분'];
                for (let i = 0; i < 7; i++) {
                    const d = new Date(ISOweekStart);
                    d.setDate(ISOweekStart.getDate() + i);
                    headerTitles.push(`${korDays[i]}(${d.getMonth() + 1}/${d.getDate()})`);
                }

                const headerRow = worksheet.addRow(headerTitles);

                // Header Styling
                headerRow.eachCell((cell) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF1F4E78' } // Dark Blue
                    };
                    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });

                // Set column widths
                worksheet.columns = [
                    { width: 5 }, { width: 12 }, { width: 15 }, { width: 25 }, { width: 25 }, { width: 10 },
                    { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }
                ];

                // Group and Fill Data
                const groups = {};
                data.forEach(p => {
                    if (!groups[p.equipment]) groups[p.equipment] = [];
                    groups[p.equipment].push(p);
                });

                const days = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

                for (const [eq, plans] of Object.entries(groups)) {
                    const activePlans = plans.filter(p => days.some(d => p[d] && String(p[d]).trim() !== ''));
                    if (activePlans.length === 0) continue;

                    // Equipment Header Row
                    const eqHeader = worksheet.addRow([`[${eq}]`]);
                    eqHeader.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
                    eqHeader.alignment = { horizontal: 'left', vertical: 'middle' };
                    worksheet.mergeCells(eqHeader.number, 1, eqHeader.number, 13);

                    // Style ONLY cells 1 through 13 (A to M)
                    for (let c = 1; c <= 13; c++) {
                        const cell = eqHeader.getCell(c);
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
                        cell.border = { top: { style: 'medium' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                    }

                    const h = holidaysMap[eq] || {};
                    const dailySums = { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
                    let totalWeeklyPlan = 0;
                    let activeDaysCount = 0;

                    days.forEach(d => {
                        if (h[d] !== 1) activeDaysCount++;
                    });

                    activePlans.forEach((plan, idx) => {
                        // Plan Row
                        const pRowValues = [idx + 1, plan.manager, plan.model, plan.partName, plan.partNo, '계획'];
                        days.forEach(d => {
                            const val = h[d] === 1 ? 'X' : (plan[d] || '');
                            pRowValues.push(val);
                            if (val !== 'X' && val !== '') {
                                const parsedVal = parseInt(val) || 0;
                                dailySums[d] += parsedVal;
                                totalWeeklyPlan += parsedVal;
                            }
                        });
                        const pRow = worksheet.addRow(pRowValues);

                        // Actual Row
                        const aRowValues = ['', '', '', '', '', '실적'];
                        days.forEach(d => {
                            aRowValues.push(h[d] === 1 ? 'X' : (plan[`${d}_act`] || ''));
                        });
                        const aRow = worksheet.addRow(aRowValues);

                        // Style Plan/Actual Rows
                        [pRow, aRow].forEach(row => {
                            row.eachCell((cell, colNum) => {
                                cell.border = {
                                    top: { style: 'thin' },
                                    left: { style: 'thin' },
                                    bottom: { style: 'thin' },
                                    right: { style: 'thin' }
                                };
                                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                                if (colNum === 6) { // '구분' column
                                    cell.font = { bold: true };
                                    if (cell.value === '계획') cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F0FF' } }; // Light Blue
                                    else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }; // light Gray
                                }
                                // Highlight completed cells
                                if (colNum > 6) {
                                    const day = days[colNum - 7];
                                    const pVal = parseInt(plan[day]) || 0;
                                    const aVal = parseInt(plan[`${day}_act`]) || 0;
                                    if (aVal >= pVal && pVal > 0) {
                                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
                                    }
                                }
                            });
                        });
                    });

                    const averagePerDay = activeDaysCount > 0 ? (totalWeeklyPlan / activeDaysCount) : 0;

                    // Totals Row
                    const tRowValues = ['', '', '', '', '', '일별 계획 합계'];
                    days.forEach(d => tRowValues.push(dailySums[d] || ''));
                    const tRow = worksheet.addRow(tRowValues);
                    tRow.eachCell((cell, colNum) => {
                        cell.font = { bold: true };

                        // Highlight Overloaded Days in Red
                        if (colNum > 6) { // Past '구분' column
                            const dayIndex = colNum - 7;
                            const d = days[dayIndex];
                            const sum = dailySums[d];
                            if (sum > averagePerDay && sum > 0) {
                                cell.font = { bold: true, color: { argb: 'FFFF0000' } }; // Pure Red
                            }
                        }

                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; // Deep Gray
                        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                }

                // Dynamic Column Width Adjustment
                worksheet.columns.forEach((column, i) => {
                    let maxColumnLength = 0;
                    column.eachCell({ includeEmpty: true }, (cell) => {
                        const cellLength = getTextWidth(cell.value);
                        if (cellLength > maxColumnLength) {
                            maxColumnLength = cellLength;
                        }
                    });
                    // Set width with minimum and slight padding
                    column.width = Math.max(8, maxColumnLength + 2);
                });

                // Send response
                res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                res.setHeader('Content-Disposition', `attachment; filename=Integrated_Plan_${weekId}.xlsx`);
                await workbook.xlsx.write(res);
                res.end();
            });
        });
    } catch (err) {
        console.error('Export styling failed:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

app.put('/api/plans-actuals', (req, res) => {
    const actuals = req.body.actuals || [];
    if (actuals.length === 0) {
        return res.json({ success: true, message: 'Nothing to update.' });
    }

    db.serialize(() => {
        db.run(`BEGIN TRANSACTION;`);
        const stmt = db.prepare(`UPDATE plans SET 
            mon_act = ?, tue_act = ?, wed_act = ?, thu_act = ?, 
            fri_act = ?, sat_act = ?, sun_act = ? 
            WHERE id = ?`);

        actuals.forEach(a => {
            stmt.run([
                a.mon_act || "", a.tue_act || "", a.wed_act || "", a.thu_act || "",
                a.fri_act || "", a.sat_act || "", a.sun_act || "",
                a.id
            ]);
        });

        stmt.finalize((err) => {
            if (err) {
                db.run(`ROLLBACK;`);
                return res.status(500).json({ success: false, error: err.message });
            }
            db.run(`COMMIT;`, (err) => {
                if (err) {
                    return res.status(500).json({ success: false, error: err.message });
                }
                res.json({ success: true, message: 'Actuals saved successfully.' });
            });
        });
    });
});

const HOST = '0.0.0.0'; // Allow external access
app.listen(PORT, HOST, () => {
    console.log(`Server is running on http://${HOST}:${PORT}`);
    console.log(`Local Access: http://localhost:${PORT}`);
    console.log(`Network Access: http://10.33.56.86:${PORT}`);
});
