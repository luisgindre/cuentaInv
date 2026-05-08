# CuentaInv

Sistema desarrollado en Node.js para procesar archivos Excel y cargar automáticamente su contenido en una tabla de Microsoft SQL Server.

El proyecto recorre todos los archivos ubicados en una carpeta configurada, interpreta la información del Excel y realiza inserciones masivas en la base de datos.

---

# ¿Qué hace el sistema?

El proceso realiza automáticamente las siguientes tareas:

1. Lee todos los archivos Excel de una carpeta local.
2. Abre la primera hoja de cada archivo.
3. Busca la fila donde aparece el texto:

```txt
Fecha imputación
```

4. Omite todas las filas anteriores a esa cabecera.
5. Convierte cada fila del Excel en un registro SQL.
6. Inserta los datos en SQL Server.
7. Procesa todos los archivos encontrados en la carpeta.

---

# Validaciones actuales

El sistema actualmente contempla las siguientes validaciones y comportamientos:

## Lectura de archivos

* Solo procesa elementos que sean archivos.
* Ignora carpetas internas.

## Excel

* Utiliza únicamente la primera hoja del archivo.
* Verifica que la hoja tenga contenido (`!ref`).
* Busca automáticamente la fila de inicio según la cabecera `Fecha imputación`.

## Datos

* Los valores vacíos se convierten en `NULL`.
* Las comillas simples `'` se escapan automáticamente para evitar errores SQL.
* Los valores numéricos se insertan como números.
* Los textos se insertan entre comillas.

## Base de datos

* La conexión se abre automáticamente antes de procesar.
* La conexión se cierra al finalizar.
* Los errores se muestran en consola.

---

# Tecnologías utilizadas

* Node.js
* MSSQL (`mssql`)
* XLSX (`xlsx`)
* dotenv

---

# Instalación

## 1. Clonar el repositorio

```bash
git clone https://github.com/luisgindre/cuentaInv.git
```

## 2. Ingresar al proyecto

```bash
cd cuentaInv
```

## 3. Instalar dependencias

```bash
npm install
```

Esto instalará automáticamente:

* mssql
* xlsx
* dotenv

---

# Configuración

## Crear archivo `.env`

Copiar el archivo de ejemplo:

### Windows

```bash
copy .env.example .env
```

### Linux / Mac

```bash
cp .env.example .env
```

---

## Configurar variables de entorno

Editar `.env`:

```env
DATA_FOLDER=./data

SQL_USER=sa
SQL_PASSWORD=password
SQL_SERVER=192.168.0.44
SQL_DATABASE=CTAINV

SQL_ENCRYPT=false

SQL_TABLE=TRANSACCIONES_2023
```

---

# Variables de entorno

| Variable       | Descripción                                    |
| -------------- | ---------------------------------------------- |
| `DATA_FOLDER`  | Carpeta donde se encuentran los archivos Excel |
| `SQL_USER`     | Usuario SQL Server                             |
| `SQL_PASSWORD` | Password SQL Server                            |
| `SQL_SERVER`   | IP o hostname del servidor                     |
| `SQL_DATABASE` | Base de datos destino                          |
| `SQL_ENCRYPT`  | Define si la conexión usa SSL                  |
| `SQL_TABLE`    | Tabla destino de inserción                     |

---

# Estructura del proyecto

```txt
cuentaInv/
│
├── data/
│   └── archivos_excel.xlsx
│
├── node_modules/
│
├── .env
├── .env.example
├── .gitignore
├── cargarPresupuesto.js
├── package.json
├── package-lock.json
└── README.md
```

---

# Cómo ejecutar el proyecto

## 1. Colocar los archivos Excel

Los archivos deben copiarse dentro de la carpeta:

```txt
/data
```

o la carpeta configurada en:

```env
DATA_FOLDER=
```

---

## 2. Ejecutar el proceso

```bash
node cargarPresupuesto.js
```

---

# Resultado esperado

Durante la ejecución se mostrará algo similar a:

```txt
######## INICIO CARGA data/archivo.xlsx
Abriendo Excel
Archivo cargado
Datos guardados correctamente.
######## FIN CARGA data/archivo.xlsx
Todos los archivos han sido procesados.
```

---

# Consideraciones importantes

## Seguridad

El archivo `.env` contiene credenciales sensibles.

NO debe subirse al repositorio.

Por eso se encuentra incluido en `.gitignore`.

---

## Performance

Actualmente:

* Se inserta fila por fila.
* No utiliza transacciones.
* No utiliza bulk insert.

Para volúmenes muy grandes podría optimizarse.

---

# Mejoras futuras sugeridas

* Uso de parámetros SQL para evitar SQL Injection.
* Bulk insert.
* Logs a archivo.
* Validaciones de tipos de datos.
* Control de archivos procesados.
* Movimiento automático de archivos procesados.
* Manejo de transacciones.


---

# Autor

Luis Gindre
