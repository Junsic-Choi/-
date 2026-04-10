const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

const nodeDbPath = path.resolve(__dirname, '../server/database.sqlite');
const d1DbPath = path.join(__dirname, '../frontend/.wrangler/state/v3/d1/miniflare-D1DatabaseObject/e1f883f4e2bd4ed9e299d645d4661e4ad4e5b7cd93d84a1b0d3eeb0dbb43d4c2.sqlite');

function checkDb(name, dbPath) {
    if (!fs.existsSync(dbPath)) {
        console.log(`[${name}] Database file not found at: ${dbPath}`);
        return Promise.resolve();
    }
    return new Promise((resolve) => {
        const db = new sqlite3.Database(dbPath, (err) => {
            if (err) {
                console.log(`[${name}] Error opening DB: ${err.message}`);
                return resolve();
            }
            db.all("SELECT weekId, COUNT(*) as count FROM plans GROUP BY weekId", [], (err, rows) => {
                if (err) {
                    console.log(`[${name}] Error querying plans: ${err.message}`);
                } else {
                    console.log(`[${name}] Weekly counts:`, rows);
                }
                db.close();
                resolve();
            });
        });
    });
}

async function run() {
    console.log("--- Database Status Check ---");
    await checkDb("Node Server SQLite", nodeDbPath);
    console.log("");
    await checkDb("Local D1 (Miniflare)", d1DbPath);
    console.log("----------------------------");
}

run();
