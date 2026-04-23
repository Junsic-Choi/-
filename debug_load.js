const check = (name) => {
    try {
        require(name);
        console.log(`${name} LOADED`);
    } catch (e) {
        console.log(`${name} LOAD FAILED: ${e.message}`);
    }
};

console.log('START');
check('express');
check('@libsql/client');
check('exceljs');
check('cors');
check('dotenv');
console.log('END');
