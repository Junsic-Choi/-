import { Hono } from 'hono';
import { handle } from 'hono/cloudflare-pages';

const app = new Hono().basePath('/api');

// 1. Get equipments list (Hardcoded or distinct from DB - let's offer a fixed list for UI)
// Based on Excel "Holiday" sheet: HSP6300, HSP8000, HM2J, AH2J, Y10T, Y15T, YBM1530
const ALL_EQUIPMENTS = [
    "HSP6300", "HSP8000 #1", "HSP8000 #2", "HM2J", "AH2J", "Y10T", "Y15T", "YBM1530"
];

app.get('/equipments', (c) => {
    return c.json({ success: true, data: ALL_EQUIPMENTS });
});

// 2. Get all distinct managers from the database
app.get('/managers', async (c) => {
    try {
        const { results } = await c.env.DB.prepare(
            `SELECT DISTINCT manager FROM plans WHERE manager IS NOT NULL AND manager != '' ORDER BY manager ASC`
        ).all();

        const managers = results.map(row => row.manager);
        return c.json({ success: true, data: managers });
    } catch (err) {
        console.error('Database error [managers]:', err.message);
        return c.json({ success: false, error: err.message }, 500);
    }
});

// 3. Get plans for a specific equipment AND weekId
app.get('/plans/:equipment/:weekId', async (c) => {
    const equipment = c.req.param('equipment');
    const weekId = c.req.param('weekId');

    try {
        // First, try to get plans for the current requested week
        const { results: rows } = await c.env.DB.prepare(
            `SELECT * FROM plans WHERE equipment = ? AND weekId = ? ORDER BY id ASC`
        ).bind(equipment, weekId).all();

        // If we found data, return it
        if (rows && rows.length > 0) {
            return c.json({ success: true, data: rows });
        }

        // If no data for this week, fetch the most recent data for this equipment
        const recentWeekRow = await c.env.DB.prepare(
            `SELECT DISTINCT weekId FROM plans WHERE equipment = ? AND weekId < ? ORDER BY weekId DESC LIMIT 1`
        ).bind(equipment, weekId).first();

        if (recentWeekRow) {
            const mostRecentWeekId = recentWeekRow.weekId;
            const { results: latestWeekRows } = await c.env.DB.prepare(
                `SELECT * FROM plans WHERE equipment = ? AND weekId = ? ORDER BY id ASC`
            ).bind(equipment, mostRecentWeekId).all();

            // Create placeholder rows, wiping out the day-specific values
            const carryoverData = latestWeekRows.map(row => ({
                ...row,
                id: undefined, // Let the frontend/save layer handle new ids
                weekId: weekId, // Assign to the newly requested week
                mon: "", tue: "", wed: "", thu: "", fri: "", sat: "", sun: "",
                mon_act: "", tue_act: "", wed_act: "", thu_act: "", fri_act: "", sat_act: "", sun_act: ""
            }));
            return c.json({ success: true, data: carryoverData });
        }

        // Absolutely no data found previously or currently
        return c.json({ success: true, data: [] });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

// 3.5 Save plans for a specific equipment AND weekId
app.post('/plans/:equipment/:weekId', async (c) => {
    const equipment = c.req.param('equipment');
    const weekId = c.req.param('weekId');
    const body = await c.req.json();
    const plans = body.plans || [];

    try {
        // D1 Batch Operations instead of traditional BEGIN TRANSACTION / ROLLBACK
        const stmts = [];

        // 1. Delete existing
        stmts.push(
            c.env.DB.prepare(`DELETE FROM plans WHERE equipment = ? AND weekId = ?`).bind(equipment, weekId)
        );

        // 2. Insert new plans
        const insertStmt = c.env.DB.prepare(`
            INSERT INTO plans 
            (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);

        for (const p of plans) {
            stmts.push(insertStmt.bind(
                equipment, weekId, p.manager || "", p.model || "", p.partName || "", p.partNo || "",
                p.mon || "", p.tue || "", p.wed || "", p.thu || "", p.fri || "", p.sat || "", p.sun || ""
            ));
        }

        await c.env.DB.batch(stmts);
        return c.json({ success: true, message: 'Plans saved successfully.' });
    } catch (err) {
        console.error("Save Error", err);
        return c.json({ success: false, error: err.message }, 500);
    }
});

// Phase 18: Holidays API
app.get('/holidays/:equipment/:weekId', async (c) => {
    const equipment = c.req.param('equipment');
    const weekId = c.req.param('weekId');
    try {
        const row = await c.env.DB.prepare(
            `SELECT * FROM equipment_holidays WHERE equipment = ? AND weekId = ?`
        ).bind(equipment, weekId).first();

        return c.json({ success: true, data: row || { mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0 } });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

app.get('/holidays-all/:weekId', async (c) => {
    const weekId = c.req.param('weekId');
    try {
        const { results: rows } = await c.env.DB.prepare(
            `SELECT * FROM equipment_holidays WHERE weekId = ?`
        ).bind(weekId).all();

        const map = {};
        rows.forEach(r => {
            map[r.equipment] = r;
        });
        return c.json({ success: true, data: map });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

app.post('/holidays', async (c) => {
    const body = await c.req.json();
    const { equipment, weekId, holidays } = body;
    const { mon, tue, wed, thu, fri, sat, sun } = holidays;

    // Check if exists because D1 doesn't fully support ON CONFLICT yet in all cases depending on schema
    const sql = `
        INSERT INTO equipment_holidays (equipment, weekId, mon, tue, wed, thu, fri, sat, sun)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(equipment, weekId) DO UPDATE SET
            mon=excluded.mon, tue=excluded.tue, wed=excluded.wed, 
            thu=excluded.thu, fri=excluded.fri, sat=excluded.sat, sun=excluded.sun
    `;
    try {
        await c.env.DB.prepare(sql).bind(
            equipment, weekId, mon || 0, tue || 0, wed || 0, thu || 0, fri || 0, sat || 0, sun || 0
        ).run();
        return c.json({ success: true, message: 'Holidays updated successfully.' });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

// 4. Get consolidated plans per weekId
app.get('/plans-consolidated/:weekId', async (c) => {
    const weekId = c.req.param('weekId');

    try {
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

        const { results: rows } = await c.env.DB.prepare(sql).bind(weekId).all();

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

        return c.json({ success: true, data: consolidatedData });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

// 5. Save actuals from consolidated view
app.put('/plans-actuals', async (c) => {
    const body = await c.req.json();
    const actuals = body.actuals || [];
    if (actuals.length === 0) {
        return c.json({ success: true, message: 'Nothing to update.' });
    }

    try {
        const stmts = [];
        const updateStmt = c.env.DB.prepare(`
            UPDATE plans SET 
            mon_act = ?, tue_act = ?, wed_act = ?, thu_act = ?, 
            fri_act = ?, sat_act = ?, sun_act = ? 
            WHERE id = ?
        `);

        for (const a of actuals) {
            stmts.push(updateStmt.bind(
                a.mon_act || "", a.tue_act || "", a.wed_act || "", a.thu_act || "",
                a.fri_act || "", a.sat_act || "", a.sun_act || "",
                a.id
            ));
        }

        await c.env.DB.batch(stmts);
        return c.json({ success: true, message: 'Actuals saved successfully.' });
    } catch (err) {
        return c.json({ success: false, error: err.message }, 500);
    }
});

// Ignore styled excel export for Cloudflare Workers temporarily, as exceljs is heavily dependent on node.js
app.get('/export-excel-styled/:weekId', async (c) => {
    return c.json({ success: false, error: "서버리스 환경에서는 엑셀 스타일 모듈(exceljs)이 아직 지원되지 않습니다. 프론트엔드 단에서의 SheetJS 백업으로 엑셀 추출을 처리해야 합니다." }, 501);
});

export const onRequest = handle(app);
