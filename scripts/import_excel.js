const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const sqlite3 = require('sqlite3').verbose();

const excelPath = path.join(__dirname, '../정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx');
const dbPath = path.join(__dirname, '../server/database.sqlite');

if (!fs.existsSync(excelPath)) {
    console.error("Excel file not found at: ", excelPath);
    process.exit(1);
}

const workbook = xlsx.readFile(excelPath);
const sheetName = '정밀가공직';

if (!workbook.SheetNames.includes(sheetName)) {
    console.error(`Sheet '${sheetName}' not found.`);
    process.exit(1);
}

const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error("Error opening database: ", err.message);
        process.exit(1);
    }
});

const sheet = workbook.Sheets[sheetName];
// Convert sheet to JSON array of arrays to handle headers easily
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// Based on visual layout from powershell earlier:
// Row 1: 장비별 주간 정삭 계획(2/15~2/21)
// Row 2: W/C | 작업자 | NO | 공관 담당자 | 기종 | 품명 | 품번 | 빈칸 | 2/15 | 2/16 ...
let currentWC = "";
const plansToInsert = [];
const weekId = '2026-W08'; // Hardcoded week representing 2/15~2/21

for (let r = 2; r < data.length; r++) {
    const row = data[r];
    if (!row || row.length === 0) continue;

    // Excel gives missing cells as undefined
    const wcRaw = row[0] ? String(row[0]).trim() : "";
    if (wcRaw && wcRaw !== currentWC) {
        currentWC = wcRaw;
    }

    if (!currentWC) continue; // Skip lines before identifying equipment

    const manager = row[3] ? String(row[3]).trim() : "";
    const model = row[4] ? String(row[4]).trim() : "";
    const partName = row[5] ? String(row[5]).trim() : "";
    const partNo = row[6] ? String(row[6]).trim() : "";

    // In our powershell output, Mon (2/15) was column index 8 (0-indexed)
    // Actually powershell output said: NO | 공관 담당자 | 기종 | 품명 | 품번 | (empty) | 2/15 | 2/16
    // Which means:
    // 0: W/C, 1: 작업자, 2: NO, 3: 담당자, 4: 기종, 5: 품명, 6: 품번, 7: 빈칸, 8: 2/15(Mon), 9: 2/16(Tue)...
    const mon = row[8] ? String(row[8]).trim() : "";
    const tue = row[9] ? String(row[9]).trim() : "";
    const wed = row[10] ? String(row[10]).trim() : "";
    const thu = row[11] ? String(row[11]).trim() : "";
    const fri = row[12] ? String(row[12]).trim() : "";
    const sat = row[13] ? String(row[13]).trim() : "";
    const sun = row[14] ? String(row[14]).trim() : "";

    // Only add if there is some valid identifier
    if (manager || model || partName || partNo || mon || tue || wed || thu || fri || sat || sun) {
        plansToInsert.push({
            equipment: currentWC,
            manager, model, partName, partNo, mon, tue, wed, thu, fri, sat, sun
        });
    }
}

console.log(`Found ${plansToInsert.length} plan rows in Excel. Starting DB Insertion...`);

db.serialize(() => {
    // Clear out '2026-W08' so we don't duplicate on multiple runs
    db.run(`DELETE FROM plans WHERE weekId = ?`, [weekId], (err) => {
        if (err) console.error("Error clearing old excel data");
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
            console.log("Successfully imported excel data into SQLite!");
        }
        db.close();
    });
});
