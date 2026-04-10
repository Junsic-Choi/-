const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.resolve(__dirname, '../server/database.sqlite');
const db = new sqlite3.Database(dbPath);

const baselineWeek = '2026-W08';
const startWeek = 9;
const endWeek = 52;

console.log(`Starting bulk synchronization from ${baselineWeek} to W09~W${endWeek}...`);

db.serialize(() => {
    // 1. Fetch baseline data
    db.all(`SELECT * FROM plans WHERE weekId = ?`, [baselineWeek], (err, baselineRows) => {
        if (err) {
            console.error('Failed to fetch baseline:', err);
            return;
        }

        if (baselineRows.length === 0) {
            console.error('No rows found in baseline week (W08). Aborting.');
            db.close();
            return;
        }

        console.log(`Baseline W08 has ${baselineRows.length} rows.`);

        db.run('BEGIN TRANSACTION');

        try {
            for (let w = startWeek; w <= endWeek; w++) {
                const targetWeek = `2026-W${String(w).padStart(2, '0')}`;

                // Delete existing data for target week
                db.run(`DELETE FROM plans WHERE weekId = ?`, [targetWeek]);

                // Insert baseline data into target week
                const stmt = db.prepare(`INSERT INTO plans 
                    (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun, mon_act, tue_act, wed_act, thu_act, fri_act, sat_act, sun_act) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

                baselineRows.forEach(row => {
                    stmt.run([
                        row.equipment, targetWeek, row.manager, row.model, row.partName, row.partNo,
                        "", "", "", "", "", "", "", // Plan values cleared
                        "", "", "", "", "", "", ""  // Actual values cleared
                    ]);
                });
                stmt.finalize();
                console.log(`Synced ${targetWeek}...`);
            }

            db.run('COMMIT', (err) => {
                if (err) {
                    console.error('Commit failed:', err);
                } else {
                    console.log('Bulk synchronization completed successfully!');
                }
                db.close();
            });
        } catch (error) {
            console.error('Sync error:', error);
            db.run('ROLLBACK');
            db.close();
        }
    });
});
