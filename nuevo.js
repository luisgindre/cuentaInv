const XLSX = require('xlsx');
const fs = require('fs');

// Ruta del archivo Excel
const filePath = 'one.xlsx';

// Cargar el archivo Excel
const workbook = XLSX.readFile(filePath);

// Obtener la primera hoja del libro
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Obtener todas las celdas de la hoja
const range = XLSX.utils.decode_range(worksheet['!ref']);

// Convertir las celdas a un objeto JSON
const jsonData = [];
for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
  const row = [];
  for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
    const cellAddress = { c: colNum, r: rowNum };
    const cellRef = XLSX.utils.encode_cell(cellAddress);
    const cell = worksheet[cellRef];
    if (cell && cell.v !== undefined) {
      row.push(cell.v);
    } else {
      row.push(null);
    }
  }
  jsonData.push(row);
}

// Omitir las primeras 9 filas
const restOfRows = jsonData.slice(9);

// Mostrar los datos en la consola
console.log('Filas restantes después de omitir las primeras 9:');
console.log(restOfRows);

// Opcional: Guardar los datos en un archivo JSON
const jsonFilePath = 'dos.json';
fs.writeFileSync(jsonFilePath, JSON.stringify(restOfRows, null, 2));
console.log(`Datos guardados en ${jsonFilePath}`);
