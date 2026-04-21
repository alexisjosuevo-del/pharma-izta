const XLSX = require('xlsx');
const wb = XLSX.readFile('CONCENTRADO ONCOLOGICO Y NUTRICIONAL 2025 2026 TOTAL.2 BASE A.xlsb');
console.log('Sheets:', wb.SheetNames);
for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, {header: 1, range: 0});
    // Print first 5 rows to see headers
    console.log(`\nSheet: ${sheetName}`);
    for (let i = 0; i < Math.min(5, data.length); i++) {
        console.log(`Row ${i}:`, data[i]);
    }
}
