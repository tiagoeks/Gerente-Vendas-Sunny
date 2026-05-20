const xlsx = require('xlsx');
const path = require('path');
const excelPath = path.resolve('c:/Users/Tiago/Desktop/Dev/Miguel/Vendas/Bando de Dados/Base_Padrão (Sell_In).xlsx');
const workbook = xlsx.readFile(excelPath);
const sheet = workbook.Sheets['Base_Padrão (Sell_In)'];
const data = xlsx.utils.sheet_to_json(sheet, { range: 0, count: 20 });
console.log(JSON.stringify(data[0], null, 2)); // Show headers and first row
const statuses = new Set();
const almoxs = new Set();
const allData = xlsx.utils.sheet_to_json(sheet);
allData.forEach(r => {
    statuses.add(r['STATUS']);
    almoxs.add(r['Almox.']);
});
console.log("Distinct Statuses:", Array.from(statuses));
console.log("Distinct Almox:", Array.from(almoxs));
console.log("Total rows in sheet:", allData.length);
