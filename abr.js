const sql = require('mssql');

// Configuración de la conexión
const config = {
  user: 'sa',
  password: 'Carmen22',
  server: '192.168.0.44', // O el servidor donde se aloja tu base de datos
  database: 'siga_test',
  options: {
    encrypt: false // Si estás utilizando Azure SQL, establece esto en true
  }
};

// Función para conectar y ejecutar una consulta
async function executeQuery() {
  try {
    await sql.connect(config);
    const result = await sql.query('SELECT * FROM personal');
    console.dir(result);
  } catch (err) {
    console.error('Error al ejecutar la consulta:', err);
  } finally {
    // Cierra la conexión después de usarla
    sql.close();
  }
}

// Ejecutar la función para conectar y ejecutar la consulta
executeQuery();