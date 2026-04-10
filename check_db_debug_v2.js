const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.resolve(__dirname, 'server/database.sqlite');
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error('Connection error:', err);
        process.exit(1);
    }
    console.log('Connected to database.');
});

db.all("SELECT name FROM sqlite_master WHERE type='table'", (err, tables) => {
    if (err) {
        console.error('Tables query error:', err);
        process.exit(1);
    }
    console.log('Tables:', tables);

    if (tables.some(t => t.name === 'plans')) {
        db.all("SELECT weekId, COUNT(*) as count FROM plans GROUP BY weekId", [], (err, rows) => {
            if (err) {
                console.error('Plans query error:', err);
                process.exit(1);
            }
            console.log('Row counts by week:', rows);
            db.close();
        });
    } else {
        console.log('Table "plans" does not exist.');
        db.close();
    }
});
