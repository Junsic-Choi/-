const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.resolve(__dirname, '../server/database.sqlite');
const db = new sqlite3.Database(dbPath);

const weeks = ['2026-W08', '2026-W09', '2026-W52'];

db.all('SELECT weekId, count(*) as count FROM plans WHERE weekId IN (?, ?, ?) GROUP BY weekId', weeks, (err, rows) => {
    if (err) {
        console.error(err);
    } else {
        console.log('Final Counts:', rows);
    }
    db.close();
});
