let fs = require('fs');

const xlsx = require('xlsx');

const { Connection, Request } = require('tedious');

// Configura la conexión a la base de datos
const config = {
  server: '192.168.0.44',
  authentication: {
    type: 'default',
    options: {
      userName: 'sa',
      password: 'Carmen22'
    }
  },
  options: {
    // Si necesitas establecer una base de datos predeterminada
    // database: 'PRESUPUESTO',
    encrypt: true // Establecer a true si estás usando cifrado SSL/TLS
  }
};

// Crear una nueva conexión a la base de datos
const connection = new Connection(config);

// Manejar eventos de conexión
connection.on('connect', (err) => {
  if (err) {
    console.error('Error al conectar:', err.message);
  } else {
    console.log('Conexión exitosa.');

    // Ejecutar una consulta
    executeStatement();
  }
});

// Función para ejecutar una consulta
function executeStatement() {
  const request = new Request('select 1', (err, rowCount) => {
    if (err) {
      console.error('Error al ejecutar la consulta:', err.message);
    } else {
      console.log(`Consulta ejecutada correctamente. Filas afectadas: ${rowCount}`);
    }

    // Cerrar la conexión después de ejecutar la consulta
    connection.close();
  });

  // Manejar eventos de resultado de la consulta
  request.on('row', (columns) => {
    columns.forEach((column) => {
      console.log(`${column.metadata.colName}: ${column.value}`);
    });
  });

  // Ejecutar la consulta
  connection.execSql(request);
}

// Establecer manejadores de eventos para otros eventos si es necesario
connection.on('end', () => {
  console.log('Conexión cerrada.');
});

// Intentar conectar
connection.connect();


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
