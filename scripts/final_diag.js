const { createClient } = require('@libsql/client');
const path = require('path');

async function check() {
    console.log('--- DIAGNOSTIC START ---');
    const dbUrl = `file:${path.resolve(__dirname, 'server', 'database.sqlite')}`;
    const client = createClient({ url: dbUrl });

    try {
        const tables = await client.execute("SELECT name FROM sqlite_master WHERE type='table'");
        console.log('Existing Tables:', tables.rows.map(r => r.name).join(', '));
        
        if (tables.rows.some(r => r.name === 'audit_logs')) {
            console.log('[SUCCESS] audit_logs table exists!');
            const head = await client.execute("SELECT * FROM audit_logs LIMIT 1");
            console.log('Audit records count:', head.rows.length);
        } else {
            console.log('[FAIL] audit_logs table NOT FOUND.');
        }

        const serverContent = require('fs').readFileSync('server/server.js', 'utf8');
        const versionLine = serverContent.match(/SERVER VERSION: [0-9.]+/);
        console.log('Server Code Version:', versionLine ? versionLine[0] : 'Unknown');
        
        const loginContent = require('fs').readFileSync('public/login.html', 'utf8');
        console.log('Login UI Username Field:', loginContent.includes('id="username"') ? 'FOUND' : 'MISSING');

    } catch (err) {
        console.error('Diagnostic error:', err);
    }
    console.log('--- DIAGNOSTIC END ---');
}

check();
