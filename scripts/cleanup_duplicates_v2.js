const { createClient } = require('@libsql/client');
const path = require('path');
require('dotenv').config();

async function cleanup() {
    const dbUrl = process.env.DATABASE_URL || `file:${path.resolve(__dirname, '../server/database.sqlite')}`;
    const dbAuthToken = process.env.DATABASE_AUTH_TOKEN;

    const client = createClient({
        url: dbUrl,
        authToken: dbAuthToken,
    });

    try {
        console.log('Searching for case-insensitive duplicates...');
        
        const duplicates = await client.execute(`
            SELECT equipment, weekId, UPPER(partNo) as upartNo, COUNT(*) as count
            FROM plans
            WHERE partNo IS NOT NULL AND partNo != ''
            GROUP BY equipment, weekId, upartNo
            HAVING count > 1
        `);

        console.log(`Found ${duplicates.rows.length} groups of duplicates.`);

        for (const group of duplicates.rows) {
            const { equipment, weekId, upartNo } = group;
            
            const rows = await client.execute({
                sql: `SELECT id, partNo FROM plans WHERE equipment = ? AND weekId = ? AND UPPER(partNo) = ? ORDER BY id ASC`,
                args: [equipment, weekId, upartNo]
            });

            const idsToDelete = rows.rows.slice(1).map(r => r.id);
            
            if (idsToDelete.length > 0) {
                console.log(`Deleting duplicates for [${equipment}] [${weekId}] [${upartNo}]: IDs ${idsToDelete.join(', ')}`);
                await client.execute({
                    sql: `DELETE FROM plans WHERE id IN (${idsToDelete.map(() => '?').join(',')})`,
                    args: idsToDelete
                });
            }
        }

        console.log('Case-insensitive cleanup completed successfully.');
    } catch (err) {
        console.error('Cleanup failed:', err);
    }
}

cleanup();
