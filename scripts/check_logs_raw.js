require('dotenv').config();
const { createClient } = require('@libsql/client');

const client = createClient({
    url: process.env.DATABASE_URL,
    authToken: process.env.DATABASE_AUTH_TOKEN,
});

async function run() {
    try {
        const result = await client.execute("SELECT *, typeof(timestamp) as ts_type FROM audit_logs ORDER BY id DESC LIMIT 20");
        console.log("Raw Audit Logs (Latest 20):");
        console.table(result.rows.map(r => ({
            id: r.id,
            user: r.username,
            ts: r.timestamp,
            type: r.ts_type
        })));
        
        const schema = await client.execute("PRAGMA table_info(audit_logs)");
        console.log("\nTable Schema:");
        console.table(schema.rows);
    } catch (e) {
        console.error(e);
    }
}

run();
