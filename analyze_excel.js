const ExcelJS = require('exceljs');

async function main() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('정삭 장비 우선 순위(HSP HM2J AH2J 10T 15T).xlsx');
    const worksheet = workbook.worksheets[0];

    for (let i = 1; i <= 3; i++) {
        console.log(`\nRow ${i}:`);
        const row = worksheet.getRow(i);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber > 10) return;
            const fgColor = cell.fill && cell.fill.fgColor ? cell.fill.fgColor.argb : "None";
            const fontColor = cell.font && cell.font.color ? cell.font.color.argb : "None";
            const bold = cell.font ? !!cell.font.bold : false;
            const fontSize = cell.font ? cell.font.size : "None";
            const hAlign = cell.alignment ? cell.alignment.horizontal : "None";
            const vAlign = cell.alignment ? cell.alignment.vertical : "None";

            let borders = [];
            if (cell.border) {
                if (cell.border.top) borders.push("top");
                if (cell.border.bottom) borders.push("bottom");
                if (cell.border.left) borders.push("left");
                if (cell.border.right) borders.push("right");
            }
            console.log(`  ${cell.address} (${cell.value}): fill=${fgColor}, font_color=${fontColor}, bold=${bold}, size=${fontSize}, align=${hAlign}/${vAlign}, borders=${borders.join('-')}`);
        });
    }
}
main().catch(console.error);
