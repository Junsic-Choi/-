require('dotenv').config();
const express = require('express');
const { createClient } = require('@libsql/client');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();

try {
    app.use(cors());
    app.use(express.json());

    const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'DNS-FMS'; 
    const ADMIN_TOKEN = process.env.ADMIN_TOKEN || 'mps_admin_secret_token_2026';

    const client = createClient({
        url: process.env.DATABASE_URL,
        authToken: process.env.DATABASE_AUTH_TOKEN,
    });

    const run = async (sql, params = []) => {
        const result = await client.execute({ sql, args: params });
        return { lastID: result.lastInsertRowid ? Number(result.lastInsertRowid) : null, changes: result.rowsAffected };
    };

    const all = async (sql, params = []) => {
        const result = await client.execute({ sql, args: params });
        return result.rows;
    };

    const get = async (sql, params = []) => {
        const result = await client.execute({ sql, args: params });
        return result.rows[0] || null;
    };

    const checkAuth = (req, res, next) => {
        // Skip auth for login
        if (req.path === '/login' || req.path === '/api/login') return next();
        
        const authHeader = req.headers.authorization;
        if (authHeader === `Bearer ${ADMIN_TOKEN}`) {
            next();
        } else {
            res.status(401).json({ success: false, error: 'Unauthorized' });
        }
    };

    app.use('/api', (req, res, next) => {
        checkAuth(req, res, next);
    });

    const logActivity = async (req, action, details) => {
        let username = req.get('X-User-Name') || 'unknown';
        try { username = decodeURIComponent(username); } catch(e) {}
        const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
        const timestamp = Date.now();
        try {
            await run(`INSERT INTO activity_logs (username, action, details, ip, timestamp) VALUES (?, ?, ?, ?, ?)`, 
                [username, action, details, ip, timestamp]);
        } catch (err) {}
    };

    app.post('/api/login', async (req, res) => {
        const { username, password } = req.body;
        const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
        const now = Date.now();
        if (password === ADMIN_PASSWORD) {
            await run(`INSERT INTO audit_logs (username, ip, status, timestamp) VALUES (?, ?, ?, ?)`, [username || 'unknown', ip, 'SUCCESS', now]);
            res.json({ success: true, token: ADMIN_TOKEN });
        } else {
            await run(`INSERT INTO audit_logs (username, ip, status, timestamp) VALUES (?, ?, ?, ?)`, [username || 'unknown', ip, 'FAILURE', now]);
            res.status(401).json({ success: false, error: 'Invalid password' });
        }
    });

    app.get('/api/logs', async (req, res) => {
        try {
            const rows = await all(`SELECT * FROM audit_logs ORDER BY timestamp DESC LIMIT 100`);
            res.json({ success: true, data: rows });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/activity-logs', async (req, res) => {
        try {
            const rows = await all(`SELECT * FROM activity_logs ORDER BY timestamp DESC LIMIT 200`);
            res.json({ success: true, data: rows });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/standard-times', async (req, res) => {
        try {
            const rows = await all(`SELECT * FROM standard_times ORDER BY equipment ASC, partNo ASC`);
            res.json({ success: true, data: rows });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.post('/api/standard-times', async (req, res) => {
        const { equipment, partNo, partName, stdTime } = req.body;
        try {
            await run(`
                INSERT INTO standard_times (equipment, partNo, partName, stdTime) 
                VALUES (?, ?, ?, ?) 
                ON CONFLICT(equipment, partNo) 
                DO UPDATE SET partName=excluded.partName, stdTime=excluded.stdTime
            `, [equipment.trim(), partNo.trim(), (partName || "").trim(), parseInt(stdTime)]);
            
            await logActivity(req, '표준시간 등록/수정', `장비: ${equipment}, 품번: ${partNo}, 시간: ${stdTime}분`);
            res.json({ success: true, message: 'Standard time saved successfully.' });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.delete('/api/standard-times/:equipment/:partNo', async (req, res) => {
        const { equipment, partNo } = req.params;
        try {
            await run(`DELETE FROM standard_times WHERE equipment = ? AND partNo = ?`, [equipment.trim(), partNo.trim()]);
            await logActivity(req, '표준시간 삭제', `장비: ${equipment}, 품번: ${partNo}`);
            res.json({ success: true, message: 'Standard time deleted successfully.' });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.post('/api/urgent-status/:id', async (req, res) => {
        const { id } = req.params;
        const { urgentStatus } = req.body;
        try {
            await run(`UPDATE plans SET urgentStatus = ? WHERE id = ?`, [urgentStatus, id]);
            await logActivity(req, '중점 항목 변경', `ID: ${id}, 상태: ${urgentStatus || '해제'}`);
            res.json({ success: true, message: 'Urgent status updated.' });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    const ALL_EQUIPMENTS = ["HSP6300", "HSP8000 #1", "HSP8000 #2", "HM2J", "AH2J", "Y10T", "Y15T", "YBM1530"];

    app.get('/api/equipments', (req, res) => {
        res.json({ success: true, data: ALL_EQUIPMENTS });
    });

    app.get('/api/managers', async (req, res) => {
        try {
            const rows = await all(`SELECT DISTINCT manager FROM plans WHERE manager IS NOT NULL AND manager != '' ORDER BY manager ASC`);
            res.json({ success: true, data: rows.map(r => r.manager) });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/holidays/:equipment/:weekId', async (req, res) => {
        const { equipment, weekId } = req.params;
        try {
            const row = await get(`SELECT * FROM equipment_holidays WHERE equipment = ? AND weekId = ?`, [equipment, weekId]);
            const holidays = row || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
            res.json({ success: true, data: holidays });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/holidays-all/:weekId', async (req, res) => {
        const { weekId } = req.params;
        try {
            const rows = await all(`SELECT * FROM equipment_holidays WHERE weekId = ?`, [weekId]);
            const map = {};
            // Pre-populate all equipments with no holidays by default
            ALL_EQUIPMENTS.forEach(eq => {
                map[eq] = { equipment: eq, weekId, mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
            });
            // Overwrite with DB values
            rows.forEach(r => {
                map[r.equipment] = r;
            });
            res.json({ success: true, data: map });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.post('/api/holidays', async (req, res) => {
        const { equipment, weekId, holidays } = req.body;
        const { mon, tue, wed, thu, fri, sat, sun } = holidays;
        try {
            await run(`INSERT INTO equipment_holidays (equipment, weekId, mon, tue, wed, thu, fri, sat, sun) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(equipment, weekId) DO UPDATE SET mon=excluded.mon, tue=excluded.tue, wed=excluded.wed, thu=excluded.thu, fri=excluded.fri, sat=excluded.sat, sun=excluded.sun`, [equipment, weekId, mon, tue, wed, thu, fri, sat, sun]);
            res.json({ success: true });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/plans/:equipment/:weekId', async (req, res) => {
        const { equipment, weekId } = req.params;
        try {
            const rows = await all(`
                SELECT p.*, s.stdTime as standardTime FROM plans p
                LEFT JOIN standard_times s ON UPPER(TRIM(p.equipment)) = UPPER(TRIM(s.equipment)) AND UPPER(TRIM(p.partNo)) = UPPER(TRIM(s.partNo))
                WHERE p.equipment = ? AND p.weekId = ?
                ORDER BY p.id ASC
            `, [equipment, weekId]);
            if (rows.length > 0) return res.json({ success: true, data: rows });
            const pastRows = await all(`
                SELECT p.*, s.stdTime as standardTime FROM plans p
                LEFT JOIN standard_times s ON UPPER(TRIM(p.equipment)) = UPPER(TRIM(s.equipment)) AND UPPER(TRIM(p.partNo)) = UPPER(TRIM(s.partNo))
                WHERE p.equipment = ? AND p.weekId <= ?
                ORDER BY p.weekId DESC LIMIT 20
            `, [equipment, weekId]);
            if (pastRows.length > 0) {
                const mostRecentWeekId = pastRows[0].weekId;
                const latestWeekRows = pastRows.filter(r => r.weekId === mostRecentWeekId);
                const carryoverData = latestWeekRows.map(row => ({ ...row, id: undefined, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" }));
                return res.json({ success: true, data: carryoverData });
            }
            res.json({ success: true, data: [] });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.post('/api/plans/:equipment/:weekId', async (req, res) => {
        const { equipment, weekId } = req.params;
        const plans = req.body.plans || [];
        try {
            const currentRows = await all(`SELECT * FROM plans WHERE equipment = ? AND weekId = ?`, [equipment, weekId]);
            
            const batch = [{ sql: `DELETE FROM plans WHERE equipment = ? AND weekId = ?`, args: [equipment, weekId] }];
            plans.forEach(p => {
                const existing = currentRows.find(cr => cr.partNo === p.partNo && cr.partName === p.partName && cr.model === p.model);
                batch.push({
                    sql: `INSERT INTO plans (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun, mon_act, tue_act, wed_act, thu_act, fri_act, sat_act, sun_act) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                    args: [equipment, weekId, p.manager || "", p.model || "", p.partName || "", p.partNo || "", p.mon || "", p.tue || "", p.wed || "", p.thu || "", p.fri || "", p.sat || "", p.sun || "", existing ? existing.mon_act : "", existing ? existing.tue_act : "", existing ? existing.wed_act : "", existing ? existing.thu_act : "", existing ? existing.fri_act : "", existing ? existing.sat_act : "", existing ? existing.sun_act : ""]
                });
            });

            // Future weeks synchronization
            const futureWeeks = await all(`SELECT DISTINCT weekId FROM plans WHERE equipment = ? AND weekId > ? ORDER BY weekId ASC`, [equipment, weekId]);
            for (const fw of futureWeeks) {
                const targetWeek = fw.weekId;
                const targetRows = await all(`SELECT * FROM plans WHERE equipment = ? AND weekId = ?`, [equipment, targetWeek]);
                batch.push({ sql: `DELETE FROM plans WHERE equipment = ? AND weekId = ?`, args: [equipment, targetWeek] });
                plans.forEach(p => {
                    const existing = targetRows.find(tr => tr.partNo === p.partNo && tr.partName === p.partName && tr.model === p.model);
                    batch.push({
                        sql: `INSERT INTO plans (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun, mon_act, tue_act, wed_act, thu_act, fri_act, sat_act, sun_act) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                        args: [equipment, targetWeek, p.manager || "", p.model || "", p.partName || "", p.partNo || "", existing ? existing.mon : "", existing ? existing.tue : "", existing ? existing.wed : "", existing ? existing.thu : "", existing ? existing.fri : "", existing ? existing.sat : "", existing ? existing.sun : "", existing ? existing.mon_act : "", existing ? existing.tue_act : "", existing ? existing.wed_act : "", existing ? existing.thu_act : "", existing ? existing.fri_act : "", existing ? existing.sat_act : "", existing ? existing.sun_act : ""]
                    });
                });
            }

            await client.batch(batch, 'write');
            await logActivity(req, '계획 수정/저장', `${equipment} (${weekId}) - ${plans.length}개 항목`);
            res.json({ success: true });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.put('/api/plans-batch-update', async (req, res) => {
        const updates = req.body.updates || [];
        if (updates.length === 0) return res.json({ success: true, message: 'Nothing to update.' });
        try {
            const batch = updates.map(u => ({
                sql: `UPDATE plans SET mon = ?, tue = ?, wed = ?, thu = ?, fri = ?, sat = ?, sun = ? WHERE id = ?`,
                args: [u.mon || "", u.tue || "", u.wed || "", u.thu || "", u.fri || "", u.sat || "", u.sun || "", u.id]
            }));
            await client.batch(batch, 'write');
            await logActivity(req, '계획 일괄 수정 (통합)', `${updates.length}개 항목 수량 업데이트`);
            res.json({ success: true, message: 'Plans updated.' });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.put('/api/plans-actuals', async (req, res) => {
        const actuals = req.body.actuals || [];
        if (actuals.length === 0) return res.json({ success: true, message: 'Nothing to update.' });
        try {
            const batch = actuals.map(a => ({
                sql: `UPDATE plans SET mon_act = ?, tue_act = ?, wed_act = ?, thu_act = ?, fri_act = ?, sat_act = ?, sun_act = ? WHERE id = ?`,
                args: [a.mon_act || "", a.tue_act || "", a.wed_act || "", a.thu_act || "", a.fri_act || "", a.sat_act || "", a.sun_act || "", a.id]
            }));
            await client.batch(batch, 'write');
            await logActivity(req, '실적 입력', `${actuals.length}개 항목 실적 업데이트`);
            res.json({ success: true, message: 'Actuals saved.' });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    app.get('/api/plans-consolidated/:weekId', async (req, res) => {
        const { weekId } = req.params;
        try {
            const sql = `
                SELECT p.*, s.stdTime as standardTime FROM plans p
                INNER JOIN (SELECT equipment, MAX(weekId) as maxWeek FROM plans WHERE weekId <= ? GROUP BY equipment) latest 
                ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek
                LEFT JOIN standard_times s ON UPPER(TRIM(p.equipment)) = UPPER(TRIM(s.equipment)) AND UPPER(TRIM(p.partNo)) = UPPER(TRIM(s.partNo))
            `;
            const rows = await all(sql, [weekId]);
            const processed = rows.map(row => row.weekId === weekId ? row : { ...row, id: undefined, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" });
            const consolidatedData = [];
            ALL_EQUIPMENTS.forEach(eq => processed.filter(d => d.equipment.trim() === eq.trim()).forEach(row => consolidatedData.push(row)));
            res.json({ success: true, data: consolidatedData });
        } catch (err) {
            res.status(500).json({ success: false, error: err.message });
        }
    });

    function getTextWidth(text) {
        if (!text) return 0;
        const str = String(text);
        let width = 0;
        for (let i = 0; i < str.length; i++) width += str.charCodeAt(i) > 255 ? 2 : 1;
        return width;
    }

    app.get('/api/export-excel-styled/:weekId', async (req, res) => {
        const { weekId } = req.params;
        try {
            const sql = `SELECT p.* FROM plans p INNER JOIN (SELECT equipment, MAX(weekId) as maxWeek FROM plans WHERE weekId <= ? GROUP BY equipment) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek`;
            const rows = await all(sql, [weekId]);
            const holidayRows = await all(`SELECT * FROM equipment_holidays WHERE weekId = ?`, [weekId]);
            const hMap = {}; holidayRows.forEach(r => hMap[r.equipment] = r);
            const data = [];
            const processed = rows.map(row => row.weekId === weekId ? row : { ...row, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" });
            ALL_EQUIPMENTS.forEach(eq => processed.filter(d => d.equipment.trim() === eq.trim()).forEach(row => data.push(row)));

            if (data.length === 0) return res.status(404).json({ success: false, message: '데이터가 없습니다.' });

            const workbook = new ExcelJS.Workbook();
            const ws = workbook.addWorksheet('통합계획');
            ws.pageSetup = { fitToPage: true, fitToWidth: 1, orientation: 'landscape', margins: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } };

            const year = parseInt(weekId.substring(0, 4));
            const week = parseInt(weekId.substring(6, 8));
            const simpleDate = new Date(year, 0, 1 + (week - 1) * 7);
            const dow = simpleDate.getDay();
            const ISOweekStart = new Date(simpleDate);
            if (dow <= 4) ISOweekStart.setDate(simpleDate.getDate() - simpleDate.getDay() + 1);
            else ISOweekStart.setDate(simpleDate.getDate() + 8 - simpleDate.getDay());
            
            ISOweekStart.setDate(ISOweekStart.getDate() - 3);

            const korDays = ['금', '토', '일', '월', '화', '수', '목'];
            const headers = ['NO', '담당자', '기종', '품명', '품번', '중요도'];
            for (let i = 0; i < 7; i++) {
                const d = new Date(ISOweekStart); d.setDate(ISOweekStart.getDate() + i);
                headers.push(`${korDays[i]}(${d.getMonth() + 1}/${d.getDate()})`);
            }

            const hRow = ws.addRow(headers);
            hRow.eachCell(c => {
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
                c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
                c.alignment = { horizontal: 'center', vertical: 'middle' };
                c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            ws.columns = headers.map((h, i) => ({ width: i < 5 ? (i === 3 || i === 4 ? 25 : 12) : 10 }));

            const groups = {}; data.forEach(p => { if (!groups[p.equipment]) groups[p.equipment] = []; groups[p.equipment].push(p); });
            const dayKeys = ['fri', 'sat', 'sun', 'mon', 'tue', 'wed', 'thu'];

            const orderedEquipment = [];
            data.forEach(plan => {
                if (!orderedEquipment.includes(plan.equipment)) {
                    orderedEquipment.push(plan.equipment);
                }
            });

            for (const eq of orderedEquipment) {
                const plans = groups[eq];
                const activePlans = plans.filter(p => dayKeys.some(d => p[d]));
                
                activePlans.sort((a, b) => {
                    const managerA = (a.manager || '').trim();
                    const managerB = (b.manager || '').trim();
                    if (managerA !== managerB) return managerA.localeCompare(managerB, 'ko');
                    const getEarliestIndex = (plan) => {
                        for (let i = 0; i < dayKeys.length; i++) {
                            const val = parseInt(plan[dayKeys[i]]) || 0;
                            if (val > 0) return i;
                        }
                        return 999;
                    };
                    return getEarliestIndex(a) - getEarliestIndex(b);
                });

                if (activePlans.length === 0) continue;

                const eqRow = ws.addRow([`[${eq}]`]);
                ws.mergeCells(eqRow.number, 1, eqRow.number, 13);
                eqRow.eachCell(c => {
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
                    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    c.alignment = { horizontal: 'left', vertical: 'middle' };
                });

                const h = hMap[eq] || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 };
                let totalPlan = 0; let activeDays = dayKeys.filter(d => h[d] !== 1).length;
                const dailySums = {}; dayKeys.forEach(d => dailySums[d] = 0);

                activePlans.forEach((p, idx) => {
                    const isUrgent = p.urgentStatus === 'URGENT';
                    const planVals = [idx + 1, p.manager, p.model, p.partName, p.partNo, isUrgent ? '*' : ''];
                    dayKeys.forEach(d => {
                        const isHoliday = h[d] === 1;
                        const pVal = isHoliday ? '' : (p[d] || '');
                        planVals.push(pVal);
                        if (!isHoliday && pVal !== '') { 
                            const v = parseInt(pVal) || 0; 
                            dailySums[d] += v; totalPlan += v; 
                        }
                    });
                    const r1 = ws.addRow(planVals);
                    r1.eachCell((c, col) => {
                        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        c.alignment = { horizontal: 'center', vertical: 'middle' };
                        if (col === 6 && isUrgent) {
                            c.font = { bold: true, color: { argb: 'FF10B981' }, size: 14 };
                        }
                        if (col > 6) {
                            c.font = { bold: true, color: { argb: 'FF1E3A8A' } };
                            if (h[dayKeys[col - 7]] === 1) {
                                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCEBEB' } };
                            }
                        }
                    });
                });

                const avg = activeDays > 0 ? totalPlan / activeDays : 0;
                const totRow = ws.addRow(['', '', '', '', '', '합계', ...dayKeys.map(d => h[d] === 1 ? '' : (dailySums[d] || ''))]);
                totRow.eachCell((c, col) => {
                    if (col >= 6) {
                        c.font = { bold: true };
                        if (col > 6) {
                            const dayKey = dayKeys[col - 7];
                            const sum = dailySums[dayKey];
                            if (sum > avg && sum > 0) c.font = { bold: true, color: { argb: 'FFFF0000' } };
                            if (h[dayKey] === 1) {
                                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCEBEB' } };
                            }
                        }
                        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
                        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' } };
                        c.alignment = { horizontal: 'center', vertical: 'middle' };
                    }
                });
            }

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename=Integrated_Plan_${weekId}.xlsx`);
            await workbook.xlsx.write(res);
            res.end();
        } catch (err) {
            console.error('Excel export error:', err);
            res.status(500).json({ success: false, error: err.message });
        }
    });

} catch (err) {
    console.error(err);
}

module.exports = app;
