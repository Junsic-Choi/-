const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const sqlite3 = require('sqlite3').verbose();

const excelPath = path.join(__dirname, '../정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx');
const d1SqlitePath = path.join(__dirname, '../frontend/.wrangler/state/v3/d1/miniflare-D1DatabaseObject/e1f883f4e2bd4ed9e299d645d4661e4ad4e5b7cd93d84a1b0d3eeb0dbb43d4c2.sqlite');

const workbook = xlsx.readFile(excelPath);
const sheetName = '정밀가공직';
const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const db = new sqlite3.Database(d1SqlitePath);

const currentWeekId = '2026-W09'; // Targeting current week
const prevWeekId = '2026-W08';

console.log("Importing for week:", currentWeekId);

let currentWC = "";
const plansToInsert = [];

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

db.serialize(() => {
    // Clear both just in case
    db.run(`DELETE FROM plans WHERE weekId = ?`, [currentWeekId]);
    db.run(`DELETE FROM plans WHERE weekId = ?`, [prevWeekId]);

    const stmt = db.prepare(`INSERT INTO plans 
        (equipment, weekId, manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

    plansToInsert.forEach(p => {
        // Insert for current week
        stmt.run([p.equipment, currentWeekId, p.manager, p.model, p.partName, p.partNo, p.mon, p.tue, p.wed, p.thu, p.fri, p.sat, p.sun]);
        // Also insert for prev week so it shows up if they switch back
        stmt.run([p.equipment, prevWeekId, p.manager, p.model, p.partName, p.partNo, p.mon, p.tue, p.wed, p.thu, p.fri, p.sat, p.sun]);
    });

    stmt.finalize(() => {
        console.log(`Successfully imported ${plansToInsert.length} per week into D1.`);
        db.close();
    });
});
