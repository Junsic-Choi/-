const { createClient } = require('@libsql/client');
const fs = require('fs');
const path = require('path');

// Live Turso Credentials from configuration
const TURSO_URL = 'libsql://mps-db-junsic-choi.aws-ap-northeast-1.turso.io';
const TURSO_TOKEN = 'eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3NzY5MTczNjIsImlkIjoiMDE5ZGI4ODYtMzMwMS03NjhhLWE4N2MtNTFkYmM2ZTViNDc1IiwicmlkIjoiOTY3ZDVlNDMtY2VhNS00MzdkLWE0MzEtNjZkYzUyZGZkNTM0In0.WGavRktwj5BAkRh0KMuwRd_zwjZMlMfk9xS7Du5AHasdxx3jz7BkFtXIaBXQn7pqVGVSEu1WbcrzfDlCU78qBg';

const localDbFile = path.resolve(__dirname, '../turso_production_backup.sqlite');
const outputSqlFile = path.resolve(__dirname, '../turso_production_backup.sql');

// Delete existing backup file if it exists to start fresh
if (fs.existsSync(localDbFile)) {
    fs.unlinkSync(localDbFile);
}

const cloudClient = createClient({ url: TURSO_URL, authToken: TURSO_TOKEN });
const localClient = createClient({ url: `file:${localDbFile}` });

async function backup() {
    console.log('Connecting to Live Turso Cloud Database...');
    
    const tables = ['plans', 'equipment_holidays', 'audit_logs', 'activity_logs'];
    
    // 1. Create schemas locally
    console.log('Initializing schemas in local SQLite database...');
    
    await localClient.execute(`CREATE TABLE IF NOT EXISTS plans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        equipment TEXT NOT NULL,
        weekId TEXT NOT NULL DEFAULT '2026-W08',
        manager TEXT,
        model TEXT,
        partName TEXT,
        partNo TEXT,
        mon TEXT, tue TEXT, wed TEXT, thu TEXT, fri TEXT, sat TEXT, sun TEXT,
        mon_act TEXT DEFAULT '', tue_act TEXT DEFAULT '', wed_act TEXT DEFAULT '',
        thu_act TEXT DEFAULT '', fri_act TEXT DEFAULT '', sat_act TEXT DEFAULT '', sun_act TEXT DEFAULT '',
        urgentStatus TEXT DEFAULT '',
        updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    await localClient.execute(`CREATE TABLE IF NOT EXISTS equipment_holidays (
        equipment TEXT NOT NULL,
        weekId TEXT NOT NULL,
        mon INTEGER DEFAULT 0, tue INTEGER DEFAULT 0, wed INTEGER DEFAULT 0,
        thu INTEGER DEFAULT 0, fri INTEGER DEFAULT 0, sat INTEGER DEFAULT 0, sun INTEGER DEFAULT 0,
        PRIMARY KEY (equipment, weekId)
    )`);

    await localClient.execute(`CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        ip TEXT,
        status TEXT,
        timestamp TEXT
    )`);

    await localClient.execute(`CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        action TEXT,
        details TEXT,
        ip TEXT,
        timestamp TEXT
    )`);

    // Indexes
    await localClient.execute(`CREATE INDEX IF NOT EXISTS idx_plans_eq_week ON plans(equipment, weekId)`);
    await localClient.execute(`CREATE INDEX IF NOT EXISTS idx_plans_weekId ON plans(weekId)`);
    await localClient.execute(`CREATE INDEX IF NOT EXISTS idx_plans_manager ON plans(manager)`);

    let sqlDump = "-- 정삭 계획 홈페이지 Live Turso DB 백업 덤프\n";
    sqlDump += "-- 생성일시: " + new Date().toISOString() + "\n\n";

    // 2. Fetch and Transfer Data
    for (const table of tables) {
        console.log(`\nFetching data from cloud table: "${table}"...`);
        const result = await cloudClient.execute(`SELECT * FROM ${table}`);
        const rows = result.rows;
        console.log(`Retrieved ${rows.length} rows from cloud.`);

        if (rows.length === 0) continue;

        const columns = Object.keys(rows[0]);
        const placeholders = columns.map(() => '?').join(',');
        const insertSql = `INSERT INTO ${table} (${columns.join(', ')}) VALUES (${placeholders})`;

        console.log(`Writing to local SQLite database in batches...`);
        
        // Batch insertion locally
        for (let i = 0; i < rows.length; i += 100) {
            const batch = rows.slice(i, i + 100);
            const queries = batch.map(row => ({
                sql: insertSql,
                args: columns.map(col => row[col])
            }));
            await localClient.batch(queries, 'write');
        }
        console.log(`Finished local transfer for "${table}".`);

        // Generate SQL statements for standard dump
        sqlDump += `-- Table: ${table}\n`;
        rows.forEach(row => {
            const values = columns.map(col => {
                const val = row[col];
                if (val === null || val === undefined) return 'NULL';
                if (typeof val === 'string') return `'${val.replace(/'/g, "''")}'`;
                return val;
            });
            sqlDump += `INSERT INTO ${table} (${columns.join(', ')}) VALUES (${values.join(', ')});\n`;
        });
        sqlDump += "\n";
    }

    // 3. Write SQL dump file
    console.log(`\nWriting SQL dump file to: ${outputSqlFile}`);
    fs.writeFileSync(outputSqlFile, sqlDump, 'utf8');

    console.log('\n=========================================');
    console.log('🎉 DATABASE BACKUP COMPLETED SUCCESSFULLY!');
    console.log(`1. SQLite DB File: ${localDbFile}`);
    console.log(`2. SQL Dump File: ${outputSqlFile}`);
    console.log('=========================================\n');
}

backup().catch(err => {
    console.error('Backup failed:', err);
    process.exit(1);
});
