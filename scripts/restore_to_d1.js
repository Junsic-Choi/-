const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const sqlite3 = require('sqlite3').verbose();

// Path to the Excel file
const excelPath = path.join(__dirname, '../정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx');

// Path to the Wrangler Local D1 SQLite file
const d1SqlitePath = path.join(__dirname, '../frontend/.wrangler/state/v3/d1/miniflare-D1DatabaseObject/e1f883f4e2bd4ed9e299d645d4661e4ad4e5b7cd93d84a1b0d3eeb0dbb43d4c2.sqlite');

if (!fs.existsSync(excelPath)) {
    console.error("Excel file not found at: ", excelPath);
    process.exit(1);
}

if (!fs.existsSync(d1SqlitePath)) {
    console.error("D1 SQLite file not found at: ", d1SqlitePath);
    process.exit(1);
}

const workbook = xlsx.readFile(excelPath);
const sheetName = '정밀가공직';

if (!workbook.SheetNames.includes(sheetName)) {
    console.error(`Sheet '${sheetName}' not found.`);
    process.exit(1);
}

const db = new sqlite3.Database(d1SqlitePath, (err) => {
    if (err) {
        console.error("Error opening database: ", err.message);
        process.exit(1);
    }
});

const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

let currentWC = "";
const plansToInsert = [];
const weekId = '2026-W08'; // Default week

for (let r = 2; r < data.length; r++) {
    const row = data[r];
    if (!row || row.length === 0) continue;

    const wcRaw = row[0] ? String(row[0]).trim() : "";
    if (wcRaw && wcRaw !== currentWC) {
        currentWC = wcRaw;
    }

    if (!currentWC) continue;

    const manager = row[3] ? String(row[3]).trim() : "";
    const model = row[4] ? String(row[4]).trim() : "";
    const partName = row[5] ? String(row[5]).trim() : "";
    const partNo = row[6] ? String(row[6]).trim() : "";

    const mon = row[8] ? String(row[8]).trim() : "";
    const tue = row[9] ? String(row[9]).trim() : "";
    const wed = row[10] ? String(row[10]).trim() : "";
    const thu = row[11] ? String(row[11]).trim() : "";
    const fri = row[12] ? String(row[12]).trim() : "";
    const sat = row[13] ? String(row[13]).trim() : "";
    const sun = row[14] ? String(row[14]).trim() : "";

    if (manager || model || partName || partNo || mon || tue || wed || thu || fri || sat || sun) {
        plansToInsert.push({
            equipment: currentWC,
            manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun
        });
    }
}

console.log(`Found ${plansToInsert.length} plan rows in Excel. Starting DB Insertion into D1...`);

db.serialize(() => {
    // Clear out '2026-W08'
    db.run(`DELETE FROM plans WHERE weekId = ?`, [weekId], (err) => {
        if (err) console.error("Error clearing old data:", err.message);
    });

    const stmt = db.prepare(`INSERT INTO plans 
        (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

    plansToInsert.forEach(p => {
        stmt.run([
            p.equipment, weekId, p.manager, p.model, p.partName, p.partNo,
            p.mon, p.tue, p.wed, p.thu, p.fri, p.sat, p.sun
        ]);
    });

    stmt.finalize((err) => {
        if (err) {
            console.error("Insertion failed:", err);
        } else {
            console.log("Successfully imported excel data into D1 Local SQLite!");
        }
        db.close();
    });
});
