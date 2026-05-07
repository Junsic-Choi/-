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
        console.log('Searching for duplicates...');
        
        // Find duplicate (equipment, weekId, partNo)
        const duplicates = await client.execute(`
            SELECT equipment, weekId, partNo, COUNT(*) as count
            FROM plans
            WHERE partNo IS NOT NULL AND partNo != ''
            GROUP BY equipment, weekId, partNo
            HAVING count > 1
        `);

        console.log(`Found ${duplicates.rows.length} groups of duplicates.`);

        for (const group of duplicates.rows) {
            const { equipment, weekId, partNo } = group;
            
            // Get all IDs for this group
            const rows = await client.execute({
                sql: `SELECT id FROM plans WHERE equipment = ? AND weekId = ? AND partNo = ? ORDER BY id ASC`,
                args: [equipment, weekId, partNo]
            });

            // Keep the first one, delete the rest
            const idsToDelete = rows.rows.slice(1).map(r => r.id);
            
            if (idsToDelete.length > 0) {
                console.log(`Deleting duplicates for [${equipment}] [${weekId}] [${partNo}]: IDs ${idsToDelete.join(', ')}`);
                await client.execute({
                    sql: `DELETE FROM plans WHERE id IN (${idsToDelete.map(() => '?').join(',')})`,
                    args: idsToDelete
                });
            }
        }

        console.log('Cleanup completed successfully.');
    } catch (err) {
        console.error('Cleanup failed:', err);
    } finally {
        // No explicit close in libsql client if using execute, but we should let it finish
    }
}

cleanup();
