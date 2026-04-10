const xlsx = require('xlsx');
const path = require('path');

const excelPath = '정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx';
const workbook = xlsx.readFile(excelPath);
const sheet = workbook.Sheets['정밀가공직'];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

console.log("Analyzing Excel Columns...");
for (let i = 0; i < Math.min(10, data.length); i++) {
    console.log(`Row ${i}:`, JSON.stringify(data[i]));
}
