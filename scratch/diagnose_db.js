
require('dotenv').config();
const { createClient } = require('@libsql/client');

const client = createClient({
    url: process.env.DATABASE_URL,
    authToken: process.env.DATABASE_AUTH_TOKEN,
});

async function diagnose() {
    console.log("Starting DB Diagnosis...");
    
    // 1. Measure count
    const startCount = Date.now();
    const countRes = await client.execute("SELECT COUNT(*) as count FROM plans");
    console.log(`Total rows in plans: ${countRes.rows[0].count} (Time: ${Date.now() - startCount}ms)`);

    // 2. Measure problematic query
    const weekId = '2026-W20'; // Example week
    const startQuery = Date.now();
    const sql = `SELECT p.* FROM plans p INNER JOIN (SELECT equipment, MAX(weekId) as maxWeek FROM plans WHERE weekId <= ? GROUP BY equipment) latest ON p.equipment = latest.equipment AND p.weekId = latest.maxWeek`;
    const rows = await client.execute({ sql, args: [weekId] });
    console.log(`Consolidated query returned ${rows.rows.length} rows (Time: ${Date.now() - startQuery}ms)`);

    // 3. Check indexes
    const indexes = await client.execute("PRAGMA index_list('plans')");
    console.log("Current indexes on 'plans':", indexes.rows);
}

diagnose().catch(console.error);
