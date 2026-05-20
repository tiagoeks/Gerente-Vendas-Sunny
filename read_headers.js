const xlsx = require('xlsx');
const fs = require('fs');

const workbook = xlsx.readFile('../Bando de Dados/Base_Padrão (Sell_In).xlsx');

const output = {
  sheetNames: workbook.SheetNames,
  sheets: {}
};

for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length > 0) {
        output.sheets[sheetName] = {
            headers: data[0] || [],
            row1: data[1] || [],
            row2: data[2] || []
        };
    } else {
        output.sheets[sheetName] = "Empty";
    }
}

fs.writeFileSync('output_utf8.json', JSON.stringify(output, null, 2), 'utf8');
console.log("Done");
