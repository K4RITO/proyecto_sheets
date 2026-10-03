# Buscador de Beneficiarios IMDEL

Aplicación de escritorio desarrollada en **Python** para consultar información de beneficiarios almacenada en distintas hojas de un documento de **Google Sheets**.

La aplicación permite buscar beneficiarios mediante **DNI**, **nombre y apellido** o **domicilio**, visualizar los programas en los que figuran y exportar los resultados de búsqueda a un archivo **CSV**.

## ✨ Características

* 🔎 Búsqueda por **DNI**.
* 👤 Búsqueda por **nombre y apellido**.
* 🏠 Búsqueda por **calle y altura**.
* 📊 Consulta de múltiples hojas de Google Sheets.
* 🔄 Actualización manual de los datos desde Google Sheets.
* 📄 Visualización de los datos encontrados dentro de la aplicación.
* 📥 Exportación de los resultados a archivos CSV.
* ⚡ Conexión a Google Sheets en segundo plano para evitar bloquear la interfaz.
* 🖥️ Aplicación de escritorio con interfaz gráfica utilizando Tkinter.

## 🗂️ Hojas consultadas

Actualmente la aplicación está configurada para consultar las siguientes hojas:

1. **INSCRIPCIÓN CUATRIMESTRAL CAPACITACION EN OFICIOS**

   * Coordinación: Capacitación laboral y empleo
   * Programa: Capacitación laboral

2. **OFICINA EMPLEO**

   * Coordinación: Capacitación laboral y empleo
   * Programa: Inserción laboral

3. **NOMINALIZACIÓN**

   * Coordinación: Capacitación laboral y empleo
   * Programa: Plan FINES

4. **DESARROLLO AGRARIO**

   * Coordinación: Desarrollo agrario
   * Programa: Huertas familiares - Entregas KIT de semillas

5. **AVICULTURA EN COMUNIDAD**

   * Coordinación: Desarrollo agrario
   * Programa: Avicultura en comunidad

6. **PRODUCTORES**

   * Coordinación: Desarrollo agrario
   * Programa: RETEP - Registro de trabajadores de la economía popular

La configuración de estas hojas se encuentra centralizada en `SHEETS_CONFIG` dentro de `main.py`.

## 🛠️ Tecnologías utilizadas

* **Python**
* **Tkinter** — interfaz gráfica.
* **gspread** — conexión y lectura de Google Sheets.
* **CSV** — exportación de resultados.
* **PyInstaller** — generación del ejecutable para Windows.

## 📁 Estructura del proyecto

```text
proyecto_sheets/
│
├── main.py
├── main.spec
├── dist/
│   ├── main.exe
│   └── Colocar credenciales acá <-
├── .gitignore
└── README.md
```

### `main.py`

Contiene toda la lógica de la aplicación:

* construcción de la interfaz gráfica;
* conexión con Google Sheets;
* carga de datos;
* búsquedas;
* validación de entradas;
* actualización de las bases;
* generación de resultados;
* exportación a CSV.

### `main.spec`

Archivo de configuración utilizado por **PyInstaller** para generar el ejecutable de la aplicación.

### `dist/`

Contiene el ejecutable generado por PyInstaller.

### `.gitignore`

Evita subir al repositorio archivos sensibles como las credenciales de Google y directorios generados durante la compilación.

## ⚙️ Requisitos

Para ejecutar el proyecto desde Python se necesita:

* Python 3
* Una cuenta/configuración de Google con acceso al documento de Google Sheets utilizado por la aplicación.
* Las credenciales de una cuenta de servicio de Google.

## 🔐 Configuración de las credenciales

La aplicación utiliza un archivo JSON de credenciales mediante:

```python
gspread.service_account(filename=CREDENTIALS_FILE)
```

El nombre esperado actualmente es:

```text
proyecto-sheets-505320-bb8ae069f096.json
```

El archivo de credenciales **no está incluido en el repositorio** por motivos de seguridad.

Debe estar disponible en el directorio desde el que se ejecuta la aplicación.

> ⚠️ **Nunca subas el archivo JSON de credenciales a GitHub.**
>
> El archivo está incluido en `.gitignore` para evitar que sea agregado accidentalmente al repositorio.

Además, la cuenta de servicio utilizada debe tener permisos sobre el documento de Google Sheets que contiene las bases consultadas.

## 📦 Instalación para desarrollo

Clonar el repositorio:

```bash
git clone https://github.com/K4RITO/proyecto_sheets.git
cd proyecto_sheets
```

Instalar las dependencias:

```bash
pip install gspread
```

Colocar el archivo de credenciales:

```text
proyecto-sheets-505320-bb8ae069f096.json
```

en el directorio correspondiente.

Después ejecutar:

```bash
python main.py
```

## 🔎 Funcionamiento

Al iniciar la aplicación se establece una conexión con Google Sheets y se cargan los datos de las hojas configuradas.

Una vez conectada, la aplicación permite realizar tres tipos de búsqueda.

### Buscar por DNI

Ingresar un número de DNI y presionar:

```text
Buscar por DNI
```

La aplicación revisará las diferentes hojas y mostrará las coincidencias encontradas, incluyendo:

* DNI
* nombre y apellido
* teléfono
* email
* domicilio
* coordinación
* programa

### Buscar por nombre y apellido

Ingresar el apellido y el nombre:

```text
Benitez Nicolas
```

La aplicación buscará coincidencias en las bases disponibles.

### Buscar por domicilio

Ingresar la calle y la altura:

```text
Cnel. Pedro Aquino 2356
```

La aplicación buscará registros asociados a ese domicilio.

## 🔄 Actualizar las bases de datos

La aplicación mantiene los datos cargados en memoria.

Para volver a obtener la información actualizada desde Google Sheets se puede utilizar:

```text
Actualizar bases de datos
```

Esto permite actualizar la información sin cerrar y volver a abrir la aplicación.

## 📥 Exportar resultados

Los resultados obtenidos mediante las búsquedas pueden exportarse utilizando:

```text
Descargar resultados
```

Los archivos se guardan actualmente en:

```text
C:\Reportes IMDEL
```

El nombre del archivo incluye la fecha y hora de generación:

```text
Registros de búsqueda de beneficiarios - DD-MM-YYYY_HH-MM-SShs.csv
```

El CSV contiene las siguientes columnas:

```text
DNI
Nombre
Apellido
Calle
Altura
Numero Telefono
Email
Coordinacion
Programa
```

## 🏗️ Generar el ejecutable

El proyecto utiliza **PyInstaller** y dispone de un archivo `main.spec` para controlar la compilación.

Instalar PyInstaller:

```bash
pip install pyinstaller
```

Generar el ejecutable:

```bash
pyinstaller main.spec
```

El ejecutable será generado dentro de:

```text
dist/
```

con el nombre:

```text
main.exe
```

La configuración actual genera una aplicación **sin consola** (`console=False`).

### Importante

El archivo de credenciales de Google **no se incluye automáticamente en el ejecutable**. El archivo JSON debe estar disponible para que la aplicación pueda conectarse a Google Sheets.

Por este motivo, al distribuir `main.exe`, también debe configurarse correctamente el acceso a las credenciales.

## 🔒 Seguridad

El proyecto trabaja con información proveniente de bases de beneficiarios, por lo que las credenciales y los datos deben manejarse con cuidado.

Buenas prácticas:

* No subir credenciales de Google al repositorio.
* No compartir públicamente archivos JSON de cuentas de servicio.
* Mantener correctamente configurados los permisos del documento de Google Sheets.
* Evitar incluir datos reales de beneficiarios en commits, ejemplos o capturas públicas.

## 🚀 Posibles mejoras

Algunas mejoras que podrían incorporarse en futuras versiones:

* Configurar la ruta de exportación desde la interfaz.
* Permitir seleccionar dinámicamente el archivo de Google Sheets.
* Separar la configuración y las credenciales del código fuente.
* Agregar logs para facilitar el diagnóstico de errores.
* Mejorar la normalización de nombres y domicilios para permitir búsquedas más flexibles.
* Incorporar manejo más detallado de errores de conexión.
* Separar la interfaz, la lógica de búsqueda y el acceso a datos en diferentes módulos.
* Crear un instalador para facilitar la distribución de la aplicación.

## 👨‍💻 Autores

* **K4RITO** -> Baragaño Matias 
* **LucaFFortin** -> Fortin Luca 
* **gonniiiiiibk** -> Beccacece Gonzalo 

Repositorio:

https://github.com/K4RITO/proyecto_sheets
