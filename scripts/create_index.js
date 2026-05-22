require('dotenv').config();
const { createClient } = require('@libsql/client');

const dbUrl = process.env.DATABASE_URL;
const dbAuthToken = process.env.DATABASE_AUTH_TOKEN;

const client = createClient({
    url: dbUrl,
    authToken: dbAuthToken,
});

async function main() {
    try {
        console.log('Connecting to Turso Cloud DB to create optimized index...');
        
        // 1. Create Index
        const result = await client.execute('CREATE INDEX IF NOT EXISTS idx_plans_weekId_equipment ON plans(weekId, equipment)');
        console.log('Index created successfully or already existed.', result);

        // 2. Verify Indexes in DB
        const indexList = await client.execute("PRAGMA index_list('plans')");
        console.log('\n--- Current indexes on plans table ---');
        console.log(indexList.rows);

    } catch (err) {
        console.error('Failed to create index:', err);
    }
}

main();
