# ConvertidorDeOrdenes

Software de escritorio portable para Windows que convierte planillas médicas (XLSX o CSV) a formato XLS (Excel 97-2003) con modelo de salida fijo de 24 columnas (A-X).

## Requisitos

- **Windows 10 o superior**
- **.NET 8 Runtime** (para ejecutar) o **.NET 8 SDK** (para compilar)
  - Descargar desde: https://dotnet.microsoft.com/download/dotnet/8.0

## Características

- ✅ Conversión de CSV (Reconfirmatorios/Reevaluaciones) a XLS
- ✅ Conversión de XLSX (Anuales/Semestrales con múltiples solapas) a XLS
- ✅ Base de datos autoalimentada de empresas (Empresas.xlsx)
- ✅ Mapeo de prestaciones configurable (PrestacionesMap.csv)
- ✅ Normalización automática de datos (localidades, provincias, prestaciones)
- ✅ Validación de campos obligatorios
- ✅ Sistema de logs detallados
- ✅ Preview de datos antes de exportar
- ✅ Exportación a XLS con 3 hojas (Hoja1 con datos, Hoja2 y Hoja3 vacías)

## Estructura del Proyecto

```
ConvertidorDeOrdenes/
├── ConvertidorDeOrdenes.sln           # Solución de Visual Studio
├── ConvertidorDeOrdenes.Core/         # Lógica de negocio (sin UI)
│   ├── Models/                        # Modelos de datos
│   │   ├── OutputRow.cs               # Fila de salida (24 columnas A-X)
│   │   ├── CompanyRecord.cs           # Registro de empresa
│   │   ├── ParseResult.cs             # Resultado del parseo
│   │   └── ValidationResult.cs        # Resultado de validación
│   ├── Parsers/                       # Analizadores de archivos
│   │   ├── CsvOrderParser.cs          # Parser para CSV
│   │   └── XlsxOrderParser.cs         # Parser para XLSX
│   └── Services/                      # Servicios de lógica
│       ├── CompanyRepositoryExcel.cs  # Gestión de base de empresas
│       ├── Normalizer.cs              # Normalización de datos
│       ├── Validator.cs               # Validación de datos
│       ├── PrestacionMapper.cs        # Mapeo de prestaciones
│       ├── XlsExporter.cs             # Exportador a XLS (NPOI)
│       └── Logger.cs                  # Sistema de logs
├── ConvertidorDeOrdenes.Desktop/      # Aplicación WinForms
│   ├── Forms/                         # Formularios de UI
│   │   ├── WizardForm.cs              # Wizard inicial de configuración
│   │   ├── MainForm.cs                # Formulario principal
│   │   └── CompanyEditDialog.cs       # Diálogos de empresa
│   └── Program.cs                     # Punto de entrada
└── README.md                          # Este archivo
```

## Instalación y Configuración

### 1. Descargar el Proyecto

Clonar o descargar el proyecto en una carpeta local.

### 2. Archivos de Configuración

Colocar los siguientes archivos en la **misma carpeta del ejecutable** (.exe):

#### **Empresas.xlsx** (Base de datos de empresas)

Archivo Excel con las siguientes columnas:

| CUIT | CIIU | Empleador | Calle | CodPostal | Localidad | Provincia | Telefono | Fax | Mail |
|------|------|-----------|-------|-----------|-----------|-----------|----------|-----|------|

- Se crea automáticamente vacío si no existe
- Se autoalimenta al cargar empresas durante el proceso

#### **PrestacionesMap.csv** (Opcional - Mapeo de prestaciones)

Archivo CSV con formato:

```csv
Origen,Destino
ACIDO T-T-MUCONICO EN ORINA,ACIDO TT MUCONICO EN ORINA
HIDROXIPIRENO EN ORINA,1-HIDROXIPIRENO EN ORINA
```

- Si no existe, las prestaciones no se mapean
- Soporta también PrestacionesMap.xlsx

### 3. Compilar el Proyecto

#### Opción A: Desde Visual Studio 2022

1. Abrir `ConvertidorDeOrdenes.sln`
2. Seleccionar configuración **Release**
3. Menú: **Build → Build Solution** (Ctrl+Shift+B)
4. El ejecutable estará en: `ConvertidorDeOrdenes.Desktop\bin\Release\net8.0-windows\ConvertidorDeOrdenes.Desktop.exe`

#### Opción B: Desde línea de comandos (dotnet CLI)

```powershell
# Navegar a la carpeta del proyecto
cd "c:\Users\Kratos\Desktop\Seres Salud\ConvertidorDeOrdenes"

# Compilar en Release
dotnet build -c Release

# El ejecutable estará en:
# ConvertidorDeOrdenes.Desktop\bin\Release\net8.0-windows\
```

### 4. Crear Versión Portable

Para crear una versión portable (todos los archivos en una carpeta):

```powershell
# Publicar aplicación autocontenida
dotnet publish ConvertidorDeOrdenes.Desktop\ConvertidorDeOrdenes.Desktop.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -o publish

# Los archivos estarán en la carpeta "publish"
```

Luego copiar `Empresas.xlsx` y `PrestacionesMap.csv` a la carpeta `publish`.

## Uso del Software

### 1. Iniciar la Aplicación

Ejecutar `ConvertidorDeOrdenes.Desktop.exe`

### 2. Wizard Inicial

Se abrirá un wizard de configuración:

1. **Tipo de carga:**
   - **Anuales/Semestrales**: Archivo XLSX con múltiples solapas
   - **Reconfirmatorios/Reevaluaciones**: Archivo CSV

2. **Frecuencia:** (Obligatorio)
   - **A**: Anual
   - **S**: Semestral
   - **R**: Reconfirmatorio

3. **Referente:** (Obligatorio)
   - Texto libre que identifica al responsable

### 3. Seleccionar y Analizar Archivo

1. Click en **"Elegir archivo..."**
2. Seleccionar el archivo de entrada (CSV o XLSX según tipo de carga)
3. Click en **"Analizar"**

El sistema:
- Parseará el archivo
- Normalizará los datos
- Validará campos obligatorios
- Mostrará preview en grilla con las 24 columnas (A-X)
- Mostrará estadísticas: filas, empresas, warnings, errores

### 4. Resolver Datos Faltantes

Si faltan datos obligatorios (especialmente CUIT del empleador):

- **Se abrirá un diálogo automáticamente** para:
  - Seleccionar empresa de la base si hay coincidencias
  - Crear nueva empresa si no existe
  - Completar datos manualmente

Campos obligatorios:
- CUIT Empleador
- Empleador
- Localidad
- Provincia
- Frecuencia
- CUIL Trabajador
- Apellido y Nombre Trabajador
- Riesgo
- Prestación
- Referente

### 5. Exportar

1. Click en **"Exportar XLS"**
2. Elegir ubicación y nombre de archivo
3. El sistema generará un archivo `.xls` con:
   - **Hoja1**: Datos (columnas A-X)
   - **Hoja2**: Vacía
   - **Hoja3**: Vacía

## Modelo de Salida (Columnas A-X)

| Col | Nombre | Obligatorio | Descripción |
|-----|--------|-------------|-------------|
| A | CuitEmpleador | ✅ | CUIT del empleador (XX-XXXXXXXX-X) |
| B | CIIU | ❌ | Código CIIU |
| C | Empleador | ✅ | Razón social del empleador |
| D | Calle | ❌ | Domicilio calle |
| E | CodPostal | ❌ | Código postal |
| F | Localidad | ✅ | Localidad (limpia, sin CP ni sufijos) |
| G | Provincia | ✅ | Provincia (normalizada) |
| H | ABMlocProv | ❌ | Alta/Baja/Modificación localidad/provincia |
| I | Telefono | ❌ | Teléfono |
| J | Fax | ❌ | Fax |
| K | Contrato | ❌ | Número de contrato |
| L | NroEstablecimiento | ❌ | Número de establecimiento |
| M | Frecuencia | ✅ | A/S/R |
| N | Cuil | ✅ | CUIL del trabajador |
| O | NroDocumento | ❌ | Número de documento |
| P | TrabajadorApellidoNombre | ✅ | Apellido y nombre del trabajador |
| Q | Riesgo | ✅ | Descripción del riesgo (máx 90 caracteres) |
| R | DescripcionRiesgo | ❌ | Descripción extendida |
| S | ABMRiesgo | ❌ | Alta/Baja/Modificación riesgo |
| T | Prestacion | ✅ | Prestación médica (sin acentos, sin "cod:") |
| U | HistoriaClinica | ❌ | Número de historia clínica |
| V | Mail | ❌ | Email |
| W | Referente | ✅ | Referente (del wizard) |
| X | DescripcionError | ❌ | Mensajes de error/validación |

## Transformaciones Automáticas

### Prestaciones
- Remueve "cod: XXX" al final
- Elimina acentos: "tórax" → "torax"
- Aplica mapeo desde PrestacionesMap.csv si existe

### Empleador/Establecimiento
- Separa formato "NRO - NOMBRE"
- Ejemplo: "2 - Pu innovations srl" → NroEstablecimiento="2", Empleador="Pu innovations srl"

### Localidad
- Limpia códigos postales: "(1605) MUNRO-B A" → "MUNRO"
- Remueve sufijos de provincia: "JUAN BAUTISTA ALBERDI BA" → "JUAN BAUTISTA ALBERDI"

### Provincia
- Normaliza variantes:
  - "BA", "B A", "BS AS" → "BUENOS AIRES"
  - "CF", "CABA" → "CAPITAL FEDERAL"
  - "CBA" → "CORDOBA"
  - etc.

### Riesgo
- Trunca a 90 caracteres si excede
- Genera warning en log

## Logs

Los logs se guardan en la carpeta **logs/** junto al ejecutable.

Formato: `log_yyyyMMdd_HHmmss.txt`

Contenido:
- Información de entrada (archivo, tipo de carga, frecuencia)
- Cantidad de filas procesadas
- Warnings (prestación sin mapeo, truncado, etc.)
- Errores (campos obligatorios faltantes)

## Errores Comunes

### "No se puede abrir Empresas.xlsx"

- **Causa**: El archivo está abierto en Excel
- **Solución**: Cerrar Excel y reintentar

### "Error leyendo archivo CSV"

- **Causa**: Encoding incorrecto o formato inválido
- **Solución**: Verificar que el CSV use encoding ISO-8859-1 (Latin-1) y delimitador coma

### "CUIT Empleador es obligatorio"

- **Causa**: No se pudo resolver el CUIT desde el archivo ni desde Empresas.xlsx
- **Solución**: Completar manualmente en el diálogo que se abre automáticamente

## Ejemplos de Archivos de Entrada

### CSV Reconfirmatorios (ejemplo)

```csv
Contrato,CUIT,Razón social,Nro. Establecimiento,Nombre establecimiento,Teléfono,CUIL,Nombre Beneficiario,Práctica solicitada,Comentarios,Localidad,Provincia,Teléfono / Celular,Email Beneficiario,Nro Agencia,Email Agencia,Teléfono Agencia
256669,30-71554420-9,"PU INNOVATIONS S.R.L.",2,"Pu innovations srl","1162474278",20-25913386-7,"BAÑULZ HERNAN DIEGO","ACIDO T-T-MUCONICO EN ORINA cod: L31","REPETIR AC TT MUCONICO","(1605) MUNRO-B A","BUENOS AIRES"," - ","","2240","belgrano@lasegunda.com.ar","1120351500"
```

### XLSX Anuales/Semestrales (estructura)

- **Hoja "Resumen"**: Referencias a solapas de detalle
- **Solapas de detalle** (1..N): Datos de empleados con columnas:
  - CUIL
  - Beneficiario / Apellido y Nombre
  - Riesgo
  - Prestación

## Soporte y Contribuciones

Para reportar problemas o sugerencias, revisar los logs en la carpeta `logs/` y contactar al equipo de desarrollo.

## Licencia

Uso interno - Seres Salud

---

**Versión**: 1.0.0  
**Fecha**: Enero 2026  
**Framework**: .NET 8 + WinForms
# SeresSalud
