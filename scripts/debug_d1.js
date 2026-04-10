const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const d1SqlitePath = 'c:/Users/i0215099/Desktop/ANTI_TEST/frontend/.wrangler/state/v3/d1/miniflare-D1DatabaseObject/e1f883f4e2bd4ed9e299d645d4661e4ad4e5b7cd93d84a1b0d3eeb0dbb43d4c2.sqlite';

const db = new sqlite3.Database(d1SqlitePath);

console.log("Querying database contents...");

db.all("SELECT weekId, equipment, COUNT(*) as count FROM plans GROUP BY weekId, equipment", [], (err, rows) => {
    if (err) {
        console.error("Error:", err.message);
        return;
    }
    console.log("Summary of data in 'plans' table:");
    console.table(rows);

    db.all("SELECT * FROM plans LIMIT 5", [], (err, rows) => {
        if (err) {
            console.error(err);
        } else {
            console.log("\nSample rows:");
            console.log(JSON.stringify(rows, null, 2));
        }
        db.close();
    });
});
