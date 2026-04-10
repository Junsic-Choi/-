const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./server/database.sqlite');

db.all("SELECT DISTINCT manager FROM plans WHERE manager IS NOT NULL AND manager != ''", [], (err, rows) => {
    if (err) {
        console.error('DB Error:', err);
    } else {
        console.log('Managers in DB:', rows.map(r => r.manager));
    }
    db.close();
});
