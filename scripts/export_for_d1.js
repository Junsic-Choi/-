const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

const localDbPath = path.resolve(__dirname, '../server/database.sqlite');
const outputSqlPath = path.resolve(__dirname, '../dist_production_data.sql');

if (!fs.existsSync(localDbPath)) {
    console.error("Local database not found!");
    process.exit(1);
}

const db = new sqlite3.Database(localDbPath);

db.all("SELECT * FROM plans", [], (err, rows) => {
    if (err) {
        console.error(err);
        return;
    }

    let sql = "-- Migration Data\n";
    rows.forEach(row => {
        const columns = Object.keys(row).filter(k => k !== 'id');
        const values = columns.map(col => {
            const val = row[col];
            if (val === null) return 'NULL';
            if (typeof val === 'string') return `'${val.replace(/'/g, "''")}'`;
            return val;
        });
        sql += `INSERT INTO plans (${columns.join(', ')}) VALUES (${values.join(', ')});\n`;
    });

    fs.writeFileSync(outputSqlPath, sql);
    console.log(`Generated ${rows.length} insert statements in ${outputSqlPath}`);
    db.close();
});
