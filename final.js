const sql = require('mssql');
const XLSX = require('xlsx');

// Configuración de la conexión a la base de datos
const config = {
  user: 'sa',
  password: 'Carmen22',
  server: '192.168.0.44',
  database: 'PRESUPUESTO',
  options: {
    encrypt: false // Si es necesario, dependiendo de la configuración de tu servidor
  }
};


// Ruta del archivo Excel
const filePath = 'one.xlsx';

// Función para cargar y guardar los datos en la tabla MSSQL
async function cargarDatos() {
  try {
    // Conectar a la base de datos
    await sql.connect(config);

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

    // Guardar los datos en la tabla MSSQL
    const table = new sql.Table('presupuesto2022'); // Reemplaza 'nombre_de_la_tabla' con el nombre real de tu tabla
    // Agregar las columnas necesarias
    table.columns.add('Vigente', sql.Int, { nullable: true });
    table.columns.add('Unidad_Ejecutora', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('ubicacion_geografica', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Tipo_Formulario', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Tipo_Contratacion', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('T_Doc_Resp', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('T_Cpte_Generador', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Subprograma', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Subparcial', sql.Int, { nullable: true });
    table.columns.add('Subjurisdiccion', sql.Int, { nullable: true });
    table.columns.add('Proyecto', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Programa', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Principal', sql.Int, { nullable: true });
    table.columns.add('Preventivo', sql.Int, { nullable: true });
    table.columns.add('Parcial', sql.Int, { nullable: true });
    table.columns.add('Pagado', sql.Int, { nullable: true });
    table.columns.add('OGESE', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Obra', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Nro_Formulario', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('N_Doc_Resp', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('N_Cpte_Generador', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Jurisdiccion', sql.Int, { nullable: true });
    table.columns.add('Inciso', sql.Int, { nullable: true });
    table.columns.add('Fuente_Financiamiento', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Finalidad_Funcion', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Fecha_transaccion', sql.VarChar(sql.MAX), { nullable: true }); // Dependiendo del formato de fecha en tu base de datos
    table.columns.add('Fecha_imputacion', sql.VarChar(sql.MAX), { nullable: true }); // Dependiendo del formato de fecha en tu base de datos
    table.columns.add('Entidad', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Ejercicio', sql.Int, { nullable: true });
    table.columns.add('Ej_Doc_Resp', sql.Int, { nullable: true });
    table.columns.add('Ej_Cpte_Generador', sql.Int, { nullable: true });
    table.columns.add('Economico', sql.Int, { nullable: true });
    table.columns.add('Devengado', sql.Int, { nullable: true });
    table.columns.add('Descripcion3', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Descripcion2', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Descripcion1', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Delegacion', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Cuit', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Crédito_Sancion', sql.Int, { nullable: true });
    table.columns.add('Compromiso', sql.Int, { nullable: true });
    table.columns.add('Beneficiario', sql.VarChar(sql.MAX), { nullable: true });
    table.columns.add('Actividad', sql.VarChar(sql.MAX), { nullable: true });
        // Añade más columnas según sea necesario

    // Insertar datos en la tabla
    restOfRows.forEach(row => {
      table.rows.add(
        row[0], 
        row[1],
        row[2], 
        row[3], 
        row[4], 
        row[5], 
        row[6], 
        row[7], 
        row[8], 
        row[9], 
        row[10], 
        row[11], 
        row[12], 
        row[13], 
        row[14], 
        row[15], 
        row[16], 
        row[17], 
        row[18], 
        row[19], 
        row[20], 
        row[21], 
        row[22], 
        row[23], 
        row[24], 
        row[25], 
        row[26], 
        row[27], 
        row[28], 
        row[29], 
        row[30], 
        row[31], 
        row[32], 
        row[33], 
        row[34], 
        row[35], 
        row[36], 
        row[37], 
        row[38], 
        row[39], 
        row[40], 
        row[41], 
      
        
        ); // Ajusta la cantidad de columnas según tu estructura de datos
    });
    console.log(table[0])

    // Crear una solicitud de cliente y ejecutar la inserción
    const request = new sql.Request();
    await request.bulk(table);

    console.log('Datos guardados en la tabla MSSQL correctamente.');
  } catch (error) {
    console.error('Error al guardar los datos en la tabla MSSQL:', error);
  } finally {
    // Cerrar la conexión después de usarla
    await sql.close();
  }
}

// Ejecutar la función para cargar y guardar los datos
cargarDatos();
