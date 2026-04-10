const ExcelJS = require('exceljs');

async function checkConditionalFormatting() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx');
        const ws = workbook.worksheets[0];

        console.log("Reading conditional formattings:");
        // The exceljs conditionalFormattings is an array of rules
        if (ws.conditionalFormattings && ws.conditionalFormattings.length > 0) {
            ws.conditionalFormattings.forEach((cf, idx) => {
                console.log(`\nRule ${idx + 1} reference: ${cf.ref}`);
                cf.rules.forEach(rule => {
                    console.log(`   Type: ${rule.type}, Operator: ${rule.operator}, Formulas: ${rule.formulae}, Style:`, rule.style);
                });
            });
        } else {
            console.log("No conditional formatting rules found on the worksheet object.");
        }
    } catch (err) {
        console.error("Error", err);
    }
}

checkConditionalFormatting();
