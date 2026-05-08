const sql = require('mssql');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const carpeta = './data';

// Configuración de la conexión a la base de datos
const config = {
  user: 'sa',
  password: 'Carmen22',
  server: '192.168.0.44',
  database: 'CTAINV',
  options: {
    encrypt: false 
  }
};

// Función para cargar y guardar los datos en la tabla MSSQL
async function cargarDatos(filePath) {
  try {
    // Conectar a la base de datos
    const pool = await sql.connect(config);

    // Cargar el archivo Excel
    console.log('Abriendo Excel');
    const workbook = XLSX.readFile(filePath);
    console.log('Archivo Cargado');

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

    // Buscar la fila que contiene "Fecha imputación" en la columna A
    let startRow = 0;
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][0] === 'Fecha imputación') {
        startRow = i + 1; // Incluir filas después de la fila con "Fecha imputación"
        break;
      }
    }

    // Omitir las filas hasta startRow
    const restOfRows = jsonData.slice(startRow);

    // Insertar datos en la tabla MSSQL excepto la última fila
    for (let i = 0; i < restOfRows.length - 1; i++) {
      const row = restOfRows[i];
      const request = pool.request();
      const values = row.map(value => {
        // Si el valor es null, devolver 'NULL'
        if (value === null) {
          return 'NULL';
        }
        // Si el valor es una cadena de texto, devolver entre comillas
        if (typeof value === 'string') {
          return `'${value.replace("'",'')}'`;
        }
        // Si el valor es un número, devolver el número
        if (!isNaN(value)) {
          return value;
        }
        // En otros casos, devolver el valor como está
        return value;
      }).join(', ');
      const insertQuery = `INSERT INTO TRANSACCIONES_2023 (Fecha_imputacion, Fecha_transaccion, Ejercicio, Tipo_Formulario, Descripcion1, Nro_Formulario, Ej_Cpte_Generador, T_Cpte_Generador, Descripcion2, N_Cpte_Generador, Tipo_Contratacion, Beneficiario, Descripcion3, Cuit, Inciso, Principal, Parcial, Subparcial, Jurisdiccion, Subjurisdiccion, Entidad, OGESE, Programa, Subprograma, Proyecto, Actividad, Obra, Unidad_Ejecutora, Fuente_Financiamiento, Finalidad_Funcion, ubicacion_geografica, Economico, T_Doc_Resp, N_Doc_Resp, Ej_Doc_Resp, Delegacion, Crédito_Sancion, Vigente, Preventivo, Compromiso, Devengado, Pagado) 
        VALUES (${values})`;
      await request.query(insertQuery);
      /* console.log('Fila insertada correctamente.'); */
    }

    console.log('Datos guardados en la tabla MSSQL correctamente.');
  } catch (error) {
    console.error('Error al guardar los datos en la tabla MSSQL:', error);
  } finally {
    // Cerrar la conexión después de usarla
    await sql.close();
  }
}

// Leer archivos de la carpeta
fs.readdir(carpeta, async (error, archivos) => {
  if (error) {
    console.error('Error al leer la carpeta:', error);
    return;
  }

  // Iterar sobre cada archivo en la carpeta
  for (const archivo of archivos) {
    const rutaArchivo = path.join(carpeta, archivo);
    // Verificar si el elemento es un archivo (no es una carpeta)
    const esArchivo = fs.statSync(rutaArchivo).isFile();
    if (esArchivo) {
      // Llamar a la función para cargar los datos pasando la ruta del archivo
      console.log('######## INICIO CARGA', rutaArchivo);
      await cargarDatos(rutaArchivo);
      console.log('######## FIN CARGA', rutaArchivo);
    }
  }

  console.log('Todos los archivos han sido procesados.');
});
