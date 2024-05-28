const fs = require('fs');
const XLSX = require('xlsx');

function replaceSpecialChars(text) {
    return text
        .replace(/�/g, 'ã')
        .replace(/�/g, 'â')
        .replace(/�/g, 'ê')
        .replace(/�/g, 'ç')
        .replace(/�/g, 'á')
        .replace(/�/g, 'é')
        .replace(/�/g, 'í')
        .replace(/�/g, 'ó')
        .replace(/�/g, 'ú')
        .replace(/�/g, 'õ')
        .replace(/�/g, 'õ')
        .replace(/�/g, 'ã');
}

function formatCSVValue(value) {
    if (value.includes(',')) {
        return `"${value}"`;
    }
    return value;
}


const workbook = XLSX.readFile('file.xlsx');

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });

const headers = Object.keys(json[0]);
let csv = headers.map(header => replaceSpecialChars(header)).join(',') + '\n';


json.forEach(row => {
    const values = headers.map(header => {
        let cellValue = String(row[header]);
        cellValue = replaceSpecialChars(cellValue);
        return formatCSVValue(cellValue);
    });
    csv += values.join(',') + '\n';
});


fs.writeFileSync('output.csv', csv);

console.log('Arquivo CSV criado com sucesso!');