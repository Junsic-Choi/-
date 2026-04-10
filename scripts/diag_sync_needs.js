const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./server/database.sqlite');

const weeks = ['2026-W08', '2026-W09', '2026-W10'];

db.all('SELECT weekId, equipment, count(*) as count FROM plans WHERE weekId IN (?, ?, ?) GROUP BY weekId, equipment', weeks, (err, rows) => {
    if (err) {
        console.error(err);
        db.close();
        return;
    }

    const summary = {};
    rows.forEach(r => {
        if (!summary[r.weekId]) summary[r.weekId] = {};
        summary[r.weekId][r.equipment] = r.count;
    });

    console.log('--- Plan Summary ---');
    console.log(JSON.stringify(summary, null, 2));

    // Check which equipment has data in W08 but not in W09
    const w08Eqs = Object.keys(summary['2026-W08'] || {});
    const w09Eqs = Object.keys(summary['2026-W09'] || {});

    const missingInW09 = w08Eqs.filter(e => !w09Eqs.includes(e));
    console.log('\n--- Missing Equipments in W09 (that exist in W08) ---');
    console.log(missingInW09);

    // Check discrepancy in counts for existing equipments
    console.log('\n--- Row Count Discrepancies (W08 vs W09) ---');
    w09Eqs.forEach(e => {
        if (summary['2026-W08'][e] !== summary['2026-W09'][e]) {
            console.log(`${e}: W08 has ${summary['2026-W08'][e]}, W09 has ${summary['2026-W09'][e]}`);
        }
    });

    db.close();
});
