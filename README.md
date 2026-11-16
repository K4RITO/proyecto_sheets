# Buscador de Beneficiarios IMDEL

Aplicación de escritorio para Windows, desarrollada en **Python** con **Tkinter**, que permite consultar en un solo lugar a los beneficiarios de los distintos programas del IMDEL, cuyos datos están repartidos en varias hojas de un documento de **Google Sheets**.

Con ella se puede saber rápidamente **en qué programas figura una persona**, buscándola por DNI, por nombre y apellido o por domicilio, y exportar los resultados a un archivo CSV.

## Índice

- [Características](#características)
- [Programas consultados](#programas-consultados)
- [Tecnologías](#tecnologías)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Requisitos](#requisitos)
- [Configuración](#configuración)
- [Instalación y ejecución en desarrollo](#instalación-y-ejecución-en-desarrollo)
- [Uso de la aplicación](#uso-de-la-aplicación)
- [Modo sin conexión y backups](#modo-sin-conexión-y-backups)
- [Exportar resultados](#exportar-resultados)
- [Generar el ejecutable](#generar-el-ejecutable)
- [Seguridad y datos personales](#seguridad-y-datos-personales)
- [Limitaciones conocidas](#limitaciones-conocidas)
- [Posibles mejoras](#posibles-mejoras)
- [Autores](#autores)

## Características

- 🔎 **Tres tipos de búsqueda:** por DNI, por apellido y nombre, y por calle y altura.
- 📊 **Varias hojas a la vez:** consulta las 6 hojas configuradas y muestra en cuáles aparece la persona.
- 🪟 **Dos vistas de resultados:** texto (ficha por persona) o tabla con scroll horizontal y vertical.
- 📥 **Exportación a CSV** de los resultados de la última búsqueda.
- 🔄 **Actualización manual** de los datos desde Google Sheets, sin cerrar la aplicación.
- 📴 **Modo sin conexión:** si no hay internet, la aplicación carga automáticamente el último backup local.
- ⚡ **Conexión en segundo plano:** la interfaz no se congela mientras se conecta o se actualiza.
- ✅ **Validación de entradas en tiempo real** (por ejemplo, el campo DNI solo acepta números).
- 🎨 **Personalización:** modo oscuro y tamaño de letra de 10 a 18 pt.

## Programas consultados

La configuración de las hojas está centralizada en la lista `SHEETS_CONFIG` de `main.py`. Actualmente se consultan:

| # | Hoja | Coordinación | Programa |
|---|------|--------------|----------|
| 1 | INSCRIPCIÓN CUATRIMESTRAL CAPACITACION EN OFICIOS | Capacitación laboral y empleo | Capacitación laboral |
| 2 | OFICINA EMPLEO | Capacitación laboral y empleo | Inserción laboral |
| 3 | NOMINALIZACIÓN | Capacitación laboral y empleo | Plan FINES |
| 4 | DESARROLLO AGRARIO | Desarrollo agrario | Huertas familiares - Entregas KIT de semillas |
| 5 | AVICULTURA EN COMUNIDAD | Desarrollo agrario | Avicultura en comunidad |
| 6 | PRODUCTORES | Desarrollo agrario | RETEP - Registro de trabajadores de la economía popular |

Las hojas se identifican por su **posición** (`worksheet_index`) dentro del documento **"Tablas IMDEL"**, no por su nombre. Si se reordenan las pestañas en Google Sheets, hay que actualizar ese índice.

### Estructura esperada de cada hoja

La aplicación lee las filas completas de cada hoja y asume este orden de columnas:

| Columna | Contenido |
|---------|-----------|
| 0 | DNI (o CUIL) |
| 1 y 2 | Nombre y apellido (el orden varía según la hoja; cada entrada de `SHEETS_CONFIG` lo contempla en su función `formatear`) |
| 3 | Calle |
| 4 | Altura |
| 5 | Teléfono |
| 6 | Email |

## Tecnologías

- **Python 3**
- **Tkinter / ttk** — interfaz gráfica
- **gspread** — conexión y lectura de Google Sheets
- **Pillow (PIL)** — carga del logo en la interfaz
- **csv** — exportación de resultados y backups
- **PyInstaller** — generación del ejecutable para Windows

## Estructura del proyecto

```text
proyecto_sheets/
├── main.py               # Toda la lógica de la aplicación
├── main.spec             # Configuración de PyInstaller
├── logo.ico              # Icono de la ventana
├── logo.png.png          # Logo mostrado en la interfaz
├── dist/
│   ├── main.exe          # Ejecutable generado
│   ├── logo.ico
│   └── logo.png.png
├── .gitignore
└── README.md
```

`main.py` contiene la clase `BuscadorApp` (interfaz, búsquedas, backups y exportación) y la configuración de las hojas (`SHEETS_CONFIG`).

## Requisitos

- **Windows** (las rutas de exportación y backup son `C:\...`).
- **Python 3** con Tkinter (incluido en la instalación estándar de Python para Windows).
- Una **cuenta de servicio de Google** con acceso al documento de Google Sheets.
- Conexión a internet para la primera carga de datos (después puede funcionar con el backup).

## Configuración

### Credenciales de Google

La aplicación se autentica con una cuenta de servicio usando:

```python
CREDENTIALS_FILE = "Credenciales.json"
gspread.service_account(filename=CREDENTIALS_FILE)
```

Pasos:

1. Crear una cuenta de servicio en Google Cloud y descargar su clave en formato JSON.
2. Renombrar el archivo a **`Credenciales.json`**.
3. Colocarlo en la **misma carpeta desde la que se ejecuta la aplicación** (junto a `main.py` o junto a `main.exe`).
4. Compartir el documento **"Tablas IMDEL"** de Google Sheets con el email de la cuenta de servicio (figura dentro del JSON, en el campo `client_email`).

> ⚠️ `Credenciales.json` está en el `.gitignore`. **Nunca lo subas al repositorio.**

### Logos

`logo.ico` y `logo.png.png` deben estar en la misma carpeta que `main.py` o `main.exe`. La aplicación los carga al iniciar para el icono de la ventana y el logo del encabezado.

### Nombre del documento

El documento de Google Sheets se busca por nombre con la constante `SHEET_NAME` (`"Tablas IMDEL"`). Si el documento se renombra, hay que actualizarla en `main.py`.

## Instalación y ejecución en desarrollo

```bash
git clone https://github.com/K4RITO/proyecto_sheets.git
cd proyecto_sheets
pip install gspread pillow
```

Colocar `Credenciales.json` en la carpeta del proyecto y ejecutar:

```bash
python main.py
```

## Uso de la aplicación

Al iniciar, la aplicación se conecta a Google Sheets en segundo plano y carga las hojas en memoria. Mientras tanto, los campos de búsqueda permanecen deshabilitados. El estado de la conexión se muestra junto al logo.

### Buscar por DNI

Ingresar el número (solo dígitos) y presionar **Buscar por DNI** o `Enter`.
Es una búsqueda **exacta**. Si en las hojas el dato está cargado como CUIL (por ejemplo `20-12345678-3`), la aplicación extrae el DNI automáticamente antes de comparar.

### Buscar por apellido y nombre

Completar **Apellido**, **Nombre** o ambos y presionar **Buscar por nombre** o `Enter`.
Es una búsqueda **parcial** y sin distinguir mayúsculas: "ben" encuentra "Benitez". Los campos solo aceptan letras y espacios.

### Buscar por domicilio

Ingresar la **calle** (obligatoria) y, opcionalmente, la **altura**.
La calle se busca de forma parcial; la altura, si se ingresa, debe coincidir exactamente. El botón se habilita recién cuando hay una calle escrita.

> Cada búsqueda limpia los otros campos para evitar mezclar criterios.

### Resultados

Cada coincidencia muestra DNI, nombre y apellido, teléfono, email, domicilio, coordinación y programa. Una misma persona puede aparecer varias veces si está en más de un programa. Debajo de los resultados se indica la cantidad de coincidencias.

### Menú

| Menú | Opción | Descripción |
|------|--------|-------------|
| **Archivo** | Descargar resultados | Exporta la última búsqueda a CSV |
| | Actualizar bases de datos | Vuelve a leer las hojas desde Google Sheets |
| | Salir | Cierra la aplicación |
| **Ver** | Tamaño de letra | 10, 12, 14, 16 o 18 pt |
| | Modo oscuro | Alterna entre tema claro y oscuro |
| | Vista de resultados | Texto o Tabla |
| **Ayuda** | Acerca de | Información de la aplicación |

En la vista de tabla, `Shift + rueda del mouse` desplaza horizontalmente.

## Modo sin conexión y backups

Cada vez que la aplicación carga o actualiza los datos desde Google Sheets, guarda un **backup local** en:

```text
C:\Reportes IMDEL\backup\
```

con el nombre `backup - DD-MM-YYYY_HH-MM-SShs.csv`. Solo se conserva **el backup más reciente**: la carpeta se vacía antes de guardar uno nuevo.

Si al iniciar no se puede conectar a Google Sheets, la aplicación:

1. Busca el backup más reciente y lo carga.
2. Pasa a **modo sin conexión** y lo indica en la barra de estado (`SIN CONEXIÓN. Usando backup: ...`).
3. Permite buscar normalmente con esos datos.

Para volver a conectar, usar **Archivo → Actualizar bases de datos**: estando en modo sin conexión, reintenta la conexión completa. Si no existe ningún backup y no hay conexión, se muestra un error de conexión.

## Exportar resultados

**Archivo → Descargar resultados** guarda la última búsqueda en:

```text
C:\Reportes IMDEL\
```

Nombre del archivo:

```text
Registros de busqueda de beneficiarios - DD-MM-YYYY_HH-MM-SShs.csv
```

Columnas: `DNI`, `Nombre`, `Apellido`, `Calle`, `Altura`, `Numero Telefono`, `Email`, `Coordinacion`, `Programa`.

## Generar el ejecutable

```bash
pip install pyinstaller
pyinstaller main.spec
```

El ejecutable se genera en `dist/main.exe` (aplicación sin consola, `console=False`).

Para distribuirlo, en la misma carpeta que `main.exe` deben estar:

- `Credenciales.json`
- `logo.ico`
- `logo.png.png`

`main.spec` no incluye estos archivos dentro del ejecutable (`datas=[]`), así que hay que copiarlos aparte.

## Seguridad y datos personales

El proyecto maneja datos personales de beneficiarios (DNI, teléfono, email, domicilio).

- No subir `Credenciales.json` ni ningún otro JSON de cuenta de servicio al repositorio.
- Compartir el Google Sheets solo con las cuentas que lo necesiten.
- Los **backups y los CSV exportados contienen datos reales en texto plano** dentro de `C:\Reportes IMDEL`. Conviene que esa carpeta esté en equipos de confianza y que los archivos exportados se eliminen cuando ya no se necesiten.
- No incluir datos reales en commits, ejemplos ni capturas públicas.

## Limitaciones conocidas

- Las rutas de backup y exportación están fijas en `C:\Reportes IMDEL` y no se pueden cambiar desde la interfaz.
- Las hojas se identifican por su posición en el documento, no por nombre.
- La búsqueda por nombre y por domicilio no normaliza tildes ni abreviaturas (por ejemplo, "Cnel." y "Coronel" se tratan como textos distintos).
- Los logos deben estar junto a la aplicación; si faltan, la aplicación puede no iniciar correctamente.

## Posibles mejoras

- Configurar las rutas de exportación y backup desde la interfaz.
- Permitir elegir el documento de Google Sheets sin modificar el código.
- Separar configuración y credenciales del código fuente.
- Normalizar nombres y domicilios (tildes, abreviaturas) para búsquedas más flexibles.
- Agregar logs y un manejo de errores de conexión más detallado.
- Separar interfaz, lógica de búsqueda y acceso a datos en módulos distintos.
- Crear un instalador para facilitar la distribución.

## Autores

- **K4RITO** — Baragaño Matias
- **LucaFFortin** — Fortin Luca
- **gonniiiiiibk** — Beccacece Gonzalo

Repositorio: https://github.com/K4RITO/proyecto_sheets
