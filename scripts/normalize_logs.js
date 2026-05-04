require('dotenv').config();
const { createClient } = require('@libsql/client');

const client = createClient({
    url: process.env.DATABASE_URL,
    authToken: process.env.DATABASE_AUTH_TOKEN,
});

async function normalize() {
    try {
        console.log("Starting timestamp normalization...");
        
        // Normalize audit_logs
        const auditRows = await client.execute("SELECT id, timestamp FROM audit_logs");
        for (const row of auditRows.rows) {
            if (typeof row.timestamp === 'string' && row.timestamp.includes('-')) {
                const epoch = new Date(row.timestamp).getTime();
                if (!isNaN(epoch)) {
                    console.log(`Normalizing audit_log id ${row.id}: ${row.timestamp} -> ${epoch}`);
                    await client.execute({
                        sql: "UPDATE audit_logs SET timestamp = ? WHERE id = ?",
                        args: [epoch, row.id]
                    });
                }
            }
        }

        // Normalize activity_logs
        const activityRows = await client.execute("SELECT id, timestamp FROM activity_logs");
        for (const row of activityRows.rows) {
            if (typeof row.timestamp === 'string' && row.timestamp.includes('-')) {
                const epoch = new Date(row.timestamp).getTime();
                if (!isNaN(epoch)) {
                    console.log(`Normalizing activity_log id ${row.id}: ${row.timestamp} -> ${epoch}`);
                    await client.execute({
                        sql: "UPDATE activity_logs SET timestamp = ? WHERE id = ?",
                        args: [epoch, row.id]
                    });
                }
            }
        }

        console.log("Normalization complete.");
    } catch (e) {
        console.error(e);
    }
}

normalize();
