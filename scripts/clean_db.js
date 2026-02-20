const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./server/database.sqlite');

db.serialize(() => {
    // 1. Remove newlines and carriage returns
    db.run("UPDATE plans SET equipment = REPLACE(REPLACE(equipment, CHAR(13), ''), CHAR(10), '')");

    // 2. Standardize HSP8000 names
    db.run("UPDATE plans SET equipment = 'HSP8000 #1' WHERE equipment LIKE '%HSP8000%#1%'");
    db.run("UPDATE plans SET equipment = 'HSP8000 #2' WHERE equipment LIKE '%HSP8000%#2%'");

    // 3. Trim all equipment names
    db.all("SELECT id, equipment FROM plans", [], (err, rows) => {
        if (err) throw err;
        rows.forEach(row => {
            const trimmed = row.equipment.trim();
            if (trimmed !== row.equipment) {
                db.run("UPDATE plans SET equipment = ? WHERE id = ?", [trimmed, row.id]);
            }
        });
        console.log("Database cleaned.");
        db.close();
    });
});
