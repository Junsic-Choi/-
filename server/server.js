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

// 2. Get plans for a specific equipment AND weekId
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
        db.all(`SELECT * FROM plans WHERE equipment = ? ORDER BY weekId DESC LIMIT 20`, [equipment], (err, pastRows) => {
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
                    mon: "",
                    tue: "",
                    wed: "",
                    thu: "",
                    fri: "",
                    sat: "",
                    sun: ""
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
app.post('/api/plans/:equipment/:weekId', (req, res) => {
    const { equipment, weekId } = req.params;
    const plans = req.body.plans; // Array of plan objects

    db.serialize(() => {
        db.run(`BEGIN TRANSACTION;`);
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

        stmt.finalize((err) => {
            if (err) {
                db.run(`ROLLBACK;`);
                return res.status(500).json({ success: false, error: err.message });
            }
            db.run(`COMMIT;`, (err) => {
                if (err) {
                    return res.status(500).json({ success: false, error: err.message });
                }
                res.json({ success: true, message: 'Plans saved successfully.' });
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

app.get('/api/holidays-all/:weekId', (req, res) => {
    const { weekId } = req.params;
    db.all(`SELECT * FROM equipment_holidays WHERE weekId = ?`, [weekId], (err, rows) => {
        if (err) return res.status(500).json({ success: false, error: err.message });
        const map = {};
        rows.forEach(r => {
            map[r.equipment] = r;
        });
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

// 4. Get consolidated plans per weekId
app.get('/api/plans-consolidated/:weekId', (req, res) => {
    const { weekId } = req.params;
    db.all(`SELECT * FROM plans WHERE 
        weekId = ? AND (
        mon != '' OR tue != '' OR wed != '' OR 
        thu != '' OR fri != '' OR sat != '' OR sun != '' 
        )
        ORDER BY equipment, id ASC`, [weekId], (err, rows) => {
        if (err) {
            res.status(500).json({ success: false, error: err.message });
        } else {
            res.json({ success: true, data: rows });
        }
    });
});

// 5. Save actuals from consolidated view
// Expects body: { actuals: [{id: 1, mon_act: '5', tue_act: '3', ...}, ...] }
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
