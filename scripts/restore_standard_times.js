const { createClient } = require('@libsql/client');

// Live Turso Credentials from export script
const TURSO_URL = 'libsql://mps-db-junsic-choi.aws-ap-northeast-1.turso.io';
const TURSO_TOKEN = 'eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3NzY5MTczNjIsImlkIjoiMDE5ZGI4ODYtMzMwMS03NjhhLWE4N2MtNTFkYmM2ZTViNDc1IiwicmlkIjoiOTY3ZDVlNDMtY2VhNS00MzdkLWE0MzEtNjZkYzUyZGZkNTM0In0.WGavRktwj5BAkRh0KMuwRd_zwjZMlMfk9xS7Du5AHasdxx3jz7BkFtXIaBXQn7pqVGVSEu1WbcrzfDlCU78qBg';

const client = createClient({ url: TURSO_URL, authToken: TURSO_TOKEN });

async function restore() {
    console.log('Connecting to Live Turso DB to check activity logs...');
    
    // 1. Fetch recent standard time registration logs
    const result = await client.execute({
        sql: `SELECT * FROM activity_logs WHERE action = '표준시간 등록/수정' ORDER BY id DESC LIMIT 50`
    });
    
    const logs = result.rows;
    console.log(`Found ${logs.length} standard time registration log entries.`);
    
    if (logs.length === 0) {
        console.log('No recent registration log entries found to restore.');
        return;
    }
    
    const restoredCount = 0;
    const batch = [];
    
    for (const log of logs) {
        const details = log.details || '';
        // Format: "장비: HSP6300, 품번: 160601-02936C, 시간: 250분"
        const regex = /장비:\s*([^,]+),\s*품번:\s*([^,]+),\s*시간:\s*(\d+)분/;
        const match = details.match(regex);
        
        if (match) {
            const equipment = match[1].trim();
            const partNo = match[2].trim();
            const stdTime = parseInt(match[3]);
            
            console.log(`\nAnalyzing log entry: "${details}"`);
            
            // Fetch partName from plans to make sure it's accurate
            const planResult = await client.execute({
                sql: `SELECT partName FROM plans WHERE partNo = ? LIMIT 1`,
                args: [partNo]
            });
            const partName = planResult.rows[0] ? planResult.rows[0].partName : 'Unknown';
            
            console.log(`↳ Extracted: W/C='${equipment}', PartNo='${partNo}', PartName='${partName}', StdTime=${stdTime}분`);
            
            // Queue restore batch operation
            batch.push({
                sql: `INSERT OR REPLACE INTO standard_times (equipment, partNo, partName, stdTime) VALUES (?, ?, ?, ?)`,
                args: [equipment, partNo, partName, stdTime]
            });
        }
    }
    
    if (batch.length > 0) {
        console.log(`\nExecuting restore transaction of ${batch.length} items to standard_times table...`);
        await client.batch(batch, 'write');
        console.log('====================================================');
        console.log(`🎉 SUCCESS: Automatically restored ${batch.length} standard times!`);
        console.log('====================================================');
    } else {
        console.log('No valid logs matched the registration pattern.');
    }
}

restore().catch(err => {
    console.error('Restore failed:', err);
    process.exit(1);
});
