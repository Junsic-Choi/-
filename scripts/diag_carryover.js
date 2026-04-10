const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./server/database.sqlite');

const ALL_EQUIPMENTS = [
    "HSP6300", "HSP8000 #1", "HSP8000 #2", "HM2J", "AH2J", "Y10T", "Y15T", "YBM1530"
];

const weekId = '2026-W09';

db.all(`SELECT * FROM plans WHERE weekId = ? ORDER BY equipment ASC, id ASC`, [weekId], (err, currentRows) => {
    if (err) throw err;
    if (currentRows && currentRows.length > 0) {
        console.log('Existing rows found for', weekId, ':', currentRows.length);
        db.close();
    } else {
        console.log('No rows for', weekId, '- checking carryover');
        db.all(`SELECT * FROM plans WHERE weekId <= ? ORDER BY weekId DESC`, [weekId], (err, allData) => {
            if (err) throw err;
            const consolidatedData = [];
            ALL_EQUIPMENTS.forEach(eq => {
                const machineData = allData.filter(d => d.equipment.trim() === eq.trim());
                if (machineData.length > 0) {
                    const latestWeekIdFound = machineData[0].weekId;
                    const latestRows = machineData.filter(d => d.weekId === latestWeekIdFound);
                    latestRows.forEach(row => {
                        consolidatedData.push({
                            equipment: row.equipment,
                            manager: row.manager,
                            weekId: row.weekId
                        });
                    });
                }
            });
            console.log('Carryover results:', consolidatedData.length, 'rows');
            console.log('Managers found:', [...new Set(consolidatedData.map(d => d.manager))]);
            db.close();
        });
    }
});
