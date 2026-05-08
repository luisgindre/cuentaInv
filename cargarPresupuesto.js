require('dotenv').config();

const sql = require('mssql');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const carpeta = process.env.DATA_FOLDER || './data';

const config = {
  user: process.env.SQL_USER,
  password: process.env.SQL_PASSWORD,
  server: process.env.SQL_SERVER,
  database: process.env.SQL_DATABASE,
  options: {
    encrypt: process.env.SQL_ENCRYPT === 'true',
    trustServerCertificate: true,
  },
};

const tablaDestino = process.env.SQL_TABLE || 'TRANSACCIONES_2023';

async function cargarDatos(filePath) {
  try {
    const pool = await sql.connect(config);

    console.log('Abriendo Excel');
    const workbook = XLSX.readFile(filePath);
    console.log('Archivo cargado');

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet['!ref']) {
      console.log('La hoja está vacía:', filePath);
      return;
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const jsonData = [];

    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      const row = [];

      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellRef = XLSX.utils.encode_cell({ c: colNum, r: rowNum });
        const cell = worksheet[cellRef];

        row.push(cell && cell.v !== undefined ? cell.v : null);
      }

      jsonData.push(row);
    }

    let startRow = 0;

    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][0] === 'Fecha imputación') {
        startRow = i + 1;
        break;
      }
    }

    const restOfRows = jsonData.slice(startRow);

    for (let i = 0; i < restOfRows.length - 1; i++) {
      const row = restOfRows[i];

      const values = row.map(value => {
        if (value === null || value === undefined) {
          return 'NULL';
        }

        if (typeof value === 'string') {
          return `'${value.replaceAll("'", "''")}'`;
        }

        if (!isNaN(value)) {
          return value;
        }

        return `'${String(value).replaceAll("'", "''")}'`;
      }).join(', ');

      const insertQuery = `
        INSERT INTO ${tablaDestino} (
          Fecha_imputacion,
          Fecha_transaccion,
          Ejercicio,
          Tipo_Formulario,
          Descripcion1,
          Nro_Formulario,
          Ej_Cpte_Generador,
          T_Cpte_Generador,
          Descripcion2,
          N_Cpte_Generador,
          Tipo_Contratacion,
          Beneficiario,
          Descripcion3,
          Cuit,
          Inciso,
          Principal,
          Parcial,
          Subparcial,
          Jurisdiccion,
          Subjurisdiccion,
          Entidad,
          OGESE,
          Programa,
          Subprograma,
          Proyecto,
          Actividad,
          Obra,
          Unidad_Ejecutora,
          Fuente_Financiamiento,
          Finalidad_Funcion,
          ubicacion_geografica,
          Economico,
          T_Doc_Resp,
          N_Doc_Resp,
          Ej_Doc_Resp,
          Delegacion,
          Crédito_Sancion,
          Vigente,
          Preventivo,
          Compromiso,
          Devengado,
          Pagado
        )
        VALUES (${values})
      `;

      await pool.request().query(insertQuery);
    }

    console.log('Datos guardados correctamente.');
  } catch (error) {
    console.error('Error al guardar los datos:', error);
  } finally {
    await sql.close();
  }
}

fs.readdir(carpeta, async (error, archivos) => {
  if (error) {
    console.error('Error al leer la carpeta:', error);
    return;
  }

  for (const archivo of archivos) {
    const rutaArchivo = path.join(carpeta, archivo);
    const esArchivo = fs.statSync(rutaArchivo).isFile();

    if (esArchivo) {
      console.log('######## INICIO CARGA', rutaArchivo);
      await cargarDatos(rutaArchivo);
      console.log('######## FIN CARGA', rutaArchivo);
    }
  }

  console.log('Todos los archivos han sido procesados.');
});