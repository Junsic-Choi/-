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
        if (req.path === '/api/login') return next();
        const authHeader = req.headers.authorization;
        if (authHeader === `Bearer ${ADMIN_TOKEN}`) {
            next();
        } else {
            res.status(401).json({ success: false, error: 'Unauthorized' });
        }
    };

    app.use('/api', (req, res, next) => {
        if (req.path === '/login') return next();
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
        const rows = await all(`SELECT * FROM audit_logs ORDER BY timestamp DESC LIMIT 100`);
        res.json({ success: true, data: rows });
    });

    app.get('/api/activity-logs', async (req, res) => {
        const rows = await all(`SELECT * FROM activity_logs ORDER BY timestamp DESC LIMIT 200`);
        res.json({ success: true, data: rows });
    });

    app.post('/api/urgent-status/:id', async (req, res) => {
        const { id } = req.params;
        const { urgentStatus } = req.body;
        await run(`UPDATE plans SET urgentStatus = ? WHERE id = ?`, [urgentStatus, id]);
        await logActivity(req, '중점 항목 변경', `ID: ${id}, 상태: ${urgentStatus || '해제'}`);
        res.json({ success: true, message: 'Urgent status updated.' });
    });

    const ALL_EQUIPMENTS = ["HSP6300", "HSP8000 #1", "HSP8000 #2", "HM2J", "AH2J", "Y10T", "Y15T", "YBM1530"];

    app.get('/api/equipments', (req, res) => {
        res.json({ success: true, data: ALL_EQUIPMENTS });
    });

    app.get('/api/managers', async (req, res) => {
        const rows = await all(`SELECT DISTINCT manager FROM plans WHERE manager IS NOT NULL AND manager != '' ORDER BY manager ASC`);
        res.json({ success: true, data: rows.map(r => r.manager) });
    });

    app.get('/api/plans/:equipment/:weekId', async (req, res) => {
        const { equipment, weekId } = req.params;
        const rows = await all(`SELECT * FROM plans WHERE equipment = ? AND weekId = ? ORDER BY id ASC`, [equipment, weekId]);
        if (rows.length > 0) return res.json({ success: true, data: rows });
        const pastRows = await all(`SELECT * FROM plans WHERE equipment = ? AND weekId <= ? ORDER BY weekId DESC LIMIT 20`, [equipment, weekId]);
        if (pastRows.length > 0) {
            const mostRecentWeekId = pastRows[0].weekId;
            const latestWeekRows = pastRows.filter(r => r.weekId === mostRecentWeekId);
            const carryoverData = latestWeekRows.map(row => ({ ...row, id: undefined, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" }));
            return res.json({ success: true, data: carryoverData });
        }
        res.json({ success: true, data: [] });
    });

    app.post('/api/plans/:equipment/:weekId', async (req, res) => {
        const { equipment, weekId } = req.params;
        const plans = req.body.plans || [];
        const batch = [{ sql: `DELETE FROM plans WHERE equipment = ? AND weekId = ?`, args: [equipment, weekId] }];
        plans.forEach(p => {
            batch.push({
                sql: `INSERT INTO plans (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                args: [equipment, weekId, p.manager || "", p.model || "", p.partName || "", p.partNo || "", p.mon || "", p.tue || "", p.wed || "", p.thu || "", p.fri || "", p.sat || "", p.sun || ""]
            });
        });
        await client.batch(batch, 'write');
        await logActivity(req, '계획 수정/저장', `${equipment} (${weekId}) - ${plans.length}개 항목`);
        res.json({ success: true });
    });

    app.get('/api/plans-consolidated/:weekId', async (req, res) => {
        const { weekId } = req.params;
        const sql = `SELECT p.* FROM plans p INNER JOIN (SELECT equipment, MAX(weekId) as maxWeek FROM plans WHERE weekId <= ? GROUP BY equipment) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek`;
        const rows = await all(sql, [weekId]);
        const processed = rows.map(row => row.weekId === weekId ? row : { ...row, id: undefined, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" });
        const consolidatedData = [];
        ALL_EQUIPMENTS.forEach(eq => processed.filter(d => d.equipment.trim() === eq.trim()).forEach(row => consolidatedData.push(row)));
        res.json({ success: true, data: consolidatedData });
    });

    app.get('/api/export-excel-styled/:weekId', async (req, res) => {
        const { weekId } = req.params;
        const sql = `SELECT p.* FROM plans p INNER JOIN (SELECT equipment, MAX(weekId) as maxWeek FROM plans WHERE weekId <= ? GROUP BY equipment) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek`;
        const rows = await all(sql, [weekId]);
        const data = rows.map(row => row.weekId === weekId ? row : { ...row, weekId, mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "", mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: "" });
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Plan');
        ws.columns = [{header:'EQ', key:'equipment'}, {header:'Model', key:'model'}, {header:'Part', key:'partName'}, {header:'Mon', key:'mon'}, {header:'Tue', key:'tue'}, {header:'Wed', key:'wed'}, {header:'Thu', key:'thu'}, {header:'Fri', key:'fri'}, {header:'Sat', key:'sat'}, {header:'Sun', key:'sun'}];
        data.forEach(d => ws.addRow(d));
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Plan_${weekId}.xlsx`);
        await workbook.xlsx.write(res);
        res.end();
    });

} catch (err) {
    console.error(err);
}

module.exports = app;
