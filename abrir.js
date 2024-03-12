let fs = require('fs');

const xlsx = require('xlsx');
const sql = require('mssql')

async function conectar() {
    try {
        // make sure that any items are correctly URL encoded in the connection string
        await sql.connect('Server=omnius,1433;Database=PRESUPUESTO;User Id=sa;Password=Carmen22;Encrypt=true')
        const result = await sql.query`select 1 `
        console.dir(result)
    } catch (err) {
        console.log(err)
        // ... error checks
    }
};

conectar();

async function eliminarFilasHojaActiva(rutaArchivo, cantidadFilas) {
  const workbook = xlsx.readFile(rutaArchivo);

  // Get active sheet name
  const activeSheetName = workbook.SheetNames[0];
  console.log(activeSheetName);

  // Access the worksheet data
  const worksheet = workbook.Sheets[activeSheetName];
  const worksheetData = worksheet['!ref']; // Get the range of data

  // Modify the worksheet data
  const worksheetRows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // Convert to an array of rows
  worksheetRows.splice(1, cantidadFilas); // Remove the specified rows
  worksheetRows.pop(); // Remove the last row

  // Reconstruct the worksheet
  const newWorksheet = xlsx.utils.json_to_sheet([], worksheetRows); // Create a new sheet from the modified data
  workbook.Sheets[activeSheetName] = newWorksheet; // Replace the original sheet

  // Save the modified workbook
  xlsx.writeFile(workbook, 'dos.xlsx');
}

eliminarFilasHojaActiva('one.xlsx', 9);
