const xlsx = require('xlsx');
const path = require('path');
const excelPath = path.resolve('c:/Users/Tiago/Desktop/Dev/Miguel/Vendas/Bando de Dados/Base_Padrão (Sell_In).xlsx');
const workbook = xlsx.readFile(excelPath);
const sheet = workbook.Sheets['Base_Padrão (Sell_In)'];
const allData = xlsx.utils.sheet_to_json(sheet);

function excelDateToISO(serial) {
    if (!serial || typeof serial !== 'number') return null;
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info.toISOString().split('T')[0];
}

const stats = {};
let aprilTotal = 0;
let aprilCount = 0;
let aprilStatus56Count = 0;
let aprilStatus56Sum = 0;

allData.forEach(row => {
    const emissao = excelDateToISO(row['Emissao']);
    const isApril2026 = emissao && emissao.startsWith('2026-04');
    
    if (isApril2026) {
        aprilTotal += parseFloat(row['Vlr.Total']) || 0;
        aprilCount++;
        
        let status = String(row['STATUS'] || '').trim();
        let almox = String(row['Almox.'] || '').trim();
        
        if ((status === '5' || status === '6') && almox !== '20') {
            aprilStatus56Count++;
            aprilStatus56Sum += parseFloat(row['Vlr.Total']) || 0;
        }
        
        const sKey = status || 'empty';
        stats[sKey] = (stats[sKey] || 0) + 1;
    }
});

console.log("April 2026 Analysis:");
console.log("Total rows in April 2026:", aprilCount);
console.log("Total sum in April 2026 (No filters):", aprilTotal);
console.log("Rows with Status 5/6 and Almox != 20:", aprilStatus56Count);
console.log("Sum with Status 5/6 and Almox != 20:", aprilStatus56Sum);
console.log("Status distribution in April 2026:", stats);
