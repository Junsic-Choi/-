const http = require('http');

const data = JSON.stringify({ urgentStatus: 'TEST' });

// Try to find a valid ID first
const optionsGet = {
    hostname: 'localhost',
    port: 3000,
    path: '/api/plans-consolidated/2026-W17',
    method: 'GET'
};

const reqGet = http.request(optionsGet, (res) => {
    let body = '';
    res.on('data', d => body += d);
    res.on('end', () => {
        try {
            const json = JSON.parse(body);
            if (json.success && json.data.length > 0) {
                const testId = json.data[0].id;
                console.log('Testing with ID:', testId);
                if (!testId) {
                    console.log('ID is undefined/null. This is the issue.');
                    return;
                }
                testPut(testId);
            } else {
                console.log('No data found for W17');
            }
        } catch (e) {
            console.log('Failed to parse GET response or no server:', e.message);
        }
    });
});
reqGet.on('error', e => console.log('GET Error:', e.message));
reqGet.end();

function testPut(id) {
    const optionsPut = {
        hostname: 'localhost',
        port: 3000,
        path: `/api/plans/${id}/urgent`,
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': data.length
        }
    };

    const reqPut = http.request(optionsPut, (res) => {
        console.log('PUT Status:', res.statusCode);
        let body = '';
        res.on('data', d => body += d);
        res.on('end', () => console.log('PUT Response:', body));
    });
    reqPut.on('error', e => console.log('PUT Error:', e.message));
    reqPut.write(data);
    reqPut.end();
}
