const cp = require('child_process');
console.log('Spawning Node...');
const res = cp.spawnSync('node', ['server/server.js'], { encoding: 'utf8', timeout: 5000 });
console.log('Status:', res.status);
console.log('Signal:', res.signal);
console.log('Stdout:', res.stdout);
console.log('Stderr:', res.stderr);
