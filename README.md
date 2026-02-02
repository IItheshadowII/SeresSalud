# 🏥 ConvertidorDeOrdenes - Seres Salud

<div align="center">

**Sistema profesional de conversión y gestión de órdenes médicas**

[![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6?logo=windows)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)](LICENSE)

Software de escritorio para Windows que automatiza la conversión de planillas médicas (XLSX o CSV) a formato XLS estandarizado con modelo de salida de 25 columnas (A-Y).

</div>

---

## 📋 Tabla de Contenidos

- [Características Principales](#-características-principales)
- [Requisitos del Sistema](#-requisitos-del-sistema)
- [Instalación](#-instalación)
- [Configuración](#-configuración)
- [Guía de Uso](#-guía-de-uso)
- [Arquitectura](#-arquitectura)
- [Modelo de Datos](#-modelo-de-datos)
- [Transformaciones Automáticas](#-transformaciones-automáticas)
- [Solución de Problemas](#-solución-de-problemas)

---

## ✨ Características Principales

### 🔄 Procesamiento de Datos
- ✅ **Conversión CSV** - Reconfirmatorios/Reevaluaciones a XLS
- ✅ **Conversión XLSX** - Anuales/Semestrales con múltiples solapas a XLS
- ✅ **Formato de salida** - 25 columnas estandarizadas (A-Y) con columna ID
- ✅ **Exportación XLS** - Compatibilidad con Excel 97-2003

### 🗄️ Gestión de Empresas
- ✅ **Base de datos integrada** - Archivo Empresas.xlsx autoalimentado
- ✅ **Auto-resolución** - Completa automáticamente datos de empresas conocidas
- ✅ **Administración visual** - CRUD completo con interfaz gráfica
- ✅ **Búsqueda inteligente** - Filtrado por CUIT, nombre, localidad
- ✅ **Backup automático** - Copia de seguridad antes de cada eliminación

### 🎯 Validación y Normalización
- ✅ **Validación estricta** - Campos obligatorios según normativa
- ✅ **Normalización automática** - Provincias, localidades, prestaciones
- ✅ **Mapeo de prestaciones** - Configuración vía PrestacionesMap.csv
- ✅ **Reglas de negocio** - Truncado de campos, limpieza de formatos

### 🖥️ Interfaz Moderna
- ✅ **Diseño profesional** - UI moderna con Segoe UI y colores corporativos
- ✅ **Selección por ART** - Soporte multi-ART (actualmente La Segunda)
- ✅ **Wizard de configuración** - Guía paso a paso
- ✅ **Preview interactivo** - Revisión y edición antes de exportar
- ✅ **Estadísticas en tiempo real** - Filas, empresas, empleados únicos
- ✅ **Sistema de logs** - Trazabilidad completa de operaciones

### 📊 Funcionalidades Avanzadas
- ✅ **Contador de empleados únicos** - Por CUIL normalizado
- ✅ **Warnings informativos** - Alertas sin bloquear el proceso
- ✅ **Gestión de errores** - Mensajes claros en columna X
- ✅ **Código postal inteligente** - Extracción desde localidad formateada

---

## 💻 Requisitos del Sistema

| Componente | Versión Mínima | Recomendado |
|-----------|----------------|-------------|
| **Sistema Operativo** | Windows 10 | Windows 11 |
| **.NET Runtime** | 8.0 | 8.0 (última) |
| **RAM** | 4 GB | 4 GB (8 GB para archivos muy grandes) |
| **Espacio en disco** | 100 MB | 500 MB |
| **Resolución** | 1280x720 | 1920x1080 |

### Descargas

- **.NET 8 Runtime**: https://dotnet.microsoft.com/download/dotnet/8.0
- **.NET 8 SDK** (desarrollo): https://dotnet.microsoft.com/download/dotnet/8.0

---

## 📦 Instalación

### Opción 1: Instalador (Recomendado)

1. Ir a **GitHub → Releases** y descargar el instalador **ConvertidorDeOrdenes-Setup.exe** (última versión)
2. Ejecutar el instalador y completar el asistente

Notas importantes:
- Los **datos del usuario** (por ejemplo la base de empresas) se guardan en `%LOCALAPPDATA%\Seres Salud\ConvertidorDeOrdenes\...` para que no se pierdan al actualizar.
- El instalador incluye instalación condicional de **WebView2 Runtime** si el equipo no lo tiene.

### Opción 2: Compilación desde Código (Desarrolladores)

#### Requisitos Previos
- Visual Studio 2022 o superior
- .NET 8 SDK instalado
- Git (opcional)

> El repo incluye `global.json` para fijar el SDK usado por `dotnet` y evitar incompatibilidades.

#### Pasos de Compilación

**Desde Visual Studio:**
```powershell
# 1. Clonar el repositorio
git clone https://github.com/IItheshadowII/SeresSalud.git
cd SeresSalud

# 2. Abrir la solución
start ConvertidorDeOrdenes.sln

# 3. En Visual Studio:
# - Seleccionar configuración "Release"
# - Menú: Build → Build Solution (Ctrl+Shift+B)
# - El ejecutable estará en:
#   ConvertidorDeOrdenes.Desktop\bin\Release\net8.0-windows\
```

**Desde línea de comandos:**
```powershell
# 1. Navegar a la carpeta del proyecto
cd SeresSalud

# 2. Restaurar dependencias y compilar
dotnet restore
dotnet build -c Release

# 3. El ejecutable estará en:
# ConvertidorDeOrdenes.Desktop\bin\Release\net8.0-windows\ConvertidorDeOrdenes.Desktop.exe
```

### Opción 3: Crear Versión Portable

Para distribuir la aplicación sin requerir .NET instalado:

```powershell
dotnet publish ConvertidorDeOrdenes.Desktop\ConvertidorDeOrdenes.Desktop.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -o portable

# La carpeta "portable" contendrá todos los archivos necesarios
```

---

## 🚀 Releases y Updates

### Auto-update (in-app)

La app se actualiza consultando **GitHub Releases** y ejecutando el **instalador** de la versión más nueva.

Cómo funciona:

1) Al iniciar, hace un chequeo automático (máximo **1 vez cada 24 hs**).
2) Consulta `https://api.github.com/repos/IItheshadowII/SeresSalud/releases/latest`.
3) Lee `tag_name` (ej: `v1.0.15`) y lo compara con la versión local.
4) Si hay update, muestra un diálogo de confirmación.
5) Si aceptás, descarga el instalador a la carpeta temporal:
    - `%TEMP%\ConvertidorDeOrdenes-Setup-<version>.exe`
6) Ejecuta el instalador y **cierra la aplicación** para completar la actualización.

Notas:
- La notificación se “recuerda” para no insistir con la misma versión.
- El estado del updater se guarda en `%LOCALAPPDATA%\Seres Salud\ConvertidorDeOrdenes\update_state.json`.
- Los datos del usuario (base de empresas, logs, etc.) se guardan en `%LOCALAPPDATA%...`, así que los updates no deberían borrarlos.
- Si la descarga/ejecución falla, se puede instalar manualmente bajando **ConvertidorDeOrdenes-Setup.exe** desde la página de Releases.

Chequeo manual:
- Menú **Ayuda → Buscar actualizaciones...** (muestra “No hay actualizaciones disponibles” si ya estás al día).

Requisitos para que funcione:
- Los releases deben estar tageados como `vX.Y.Z`.
- El release debe incluir un asset instalador. Se prioriza el nombre exacto: **ConvertidorDeOrdenes-Setup.exe**.
- Si el repo es privado o el entorno corporativo bloquea GitHub, configurar un token:
    - Variable de entorno `SERESSALUD_GITHUB_TOKEN` (o `GITHUB_TOKEN`) con permisos para leer releases.

### Release con 1 comando (local)

Este script automatiza todo:

```powershell
pwsh -ExecutionPolicy Bypass -File .\scripts\release.ps1 -Version 1.2.3
```

Hace:
1) actualiza versión en el csproj
2) publica (`scripts/publish.ps1`)
3) genera instalador (`scripts/build-installer.ps1` → `Installer/Output/ConvertidorDeOrdenes-Setup.exe`)
4) commit + tag `v1.2.3` + push

### CI (GitHub Actions)

El workflow de release (por tags `v*`) genera el instalador y crea el GitHub Release adjuntando el setup.

Nota: el instalador offline de **WebView2 Runtime** no se versiona en git (está ignorado) y el workflow lo descarga durante el build.

---

## ⚙️ Configuración

### Archivos de Configuración

Colocar los siguientes archivos en la **misma carpeta del ejecutable** (.exe):

#### 1️⃣ **Empresas.xlsx** (Base de Datos de Empresas)

**Ubicación**: Misma carpeta que el .exe (la aplicación busca hacia arriba en las carpetas superiores y elige la copia más probable si hay varias)

**Formato** (ejemplo totalmente ficticio):

| CUIT          | CIIU | Empleador            | Calle              | CodPostal | Localidad   | Provincia   | Telefono   | Fax | Mail                     |
|---------------|------|----------------------|--------------------|-----------|-------------|-------------|------------|-----|--------------------------|
| 30-00000000-0 | 6200 | EMPRESA DEMO S.R.L.  | Calle Ficticia 123 | 1000      | CIUDAD DEMO | PROVINCIA X | 1100000000 |     | contacto@demo-local.test |

**Características**:
- ✅ Se crea automáticamente vacío si no existe
- ✅ Se autoalimenta al procesar nuevos archivos
- ✅ Soporta múltiples formatos de entrada
- ✅ Búsqueda por CUIT o nombre de empresa
- ✅ Gestión visual desde menú Empresas → Administrar

#### 2️⃣ **PrestacionesMap.csv** (Opcional - Mapeo de Prestaciones)

**Ubicación**: Misma carpeta que el .exe

**Formato CSV** (ejemplo genérico):
```csv
Origen,Destino
ACIDO T-T-MUCONICO EN ORINA,ACIDO TT MUCONICO EN ORINA
HIDROXIPIRENO EN ORINA,1-HIDROXIPIRENO EN ORINA
EXAMEN AUDIOMETRICO,AUDIOMETRIA
RX DE TORAX DE FRENTE,RX TORAX FRENTE
```

**Formato XLSX** (alternativo):
| Origen | Destino |
|--------|---------|
| ACIDO T-T-MUCONICO EN ORINA | ACIDO TT MUCONICO EN ORINA |

**Características**:
- ✅ Acepta CSV o XLSX
- ✅ Si no existe, las prestaciones no se mapean
- ✅ Se aplica después de limpiar códigos y acentos

---

## 📖 Guía de Uso

### Flujo de Trabajo Completo

```
┌─────────────────┐
│ 1. Seleccionar  │
│    ART          │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ 2. Configurar   │
│    Tipo + Freq  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ 3. Elegir       │
│    Archivo      │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ 4. Analizar     │
│    y Validar    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ 5. Revisar      │
│    Empresas     │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ 6. Exportar XLS │
└─────────────────┘
```

### 1️⃣ Selección de ART

Al iniciar la aplicación:
- Seleccionar **La Segunda** (actualmente única ART soportada)
- Click en **Siguiente**

> 💡 **Futuro**: Se agregarán más ARTs con configuraciones específicas

### 2️⃣ Configuración de Conversión

**Tipo de carga:**
- 🔵 **Anuales/Semestrales** → Archivo XLSX con múltiples solapas
- 🔵 **Reconfirmatorios/Reevaluaciones** → Archivo CSV

**Frecuencia:** (automática según tipo)
- **A** - Anual / Reconfirmatorio
- **S** - Semestral

> ⚠️ **Nota**: Referente ya no se solicita al inicio (se deja vacío)

### 3️⃣ Seleccionar Archivo

1. Click en **"Elegir archivo..."**
2. Navegar hasta el archivo de entrada:
   - **CSV** para Reconfirmatorios
   - **XLSX** para Anuales/Semestrales

### 4️⃣ Análisis y Validación

El sistema automáticamente:
- ✅ Parsea el archivo según el formato
- ✅ Extrae datos de empresa y trabajadores
- ✅ Busca empresas en Empresas.xlsx
- ✅ Autocompleta datos conocidos
- ✅ Normaliza provincias, localidades, prestaciones
- ✅ Valida campos obligatorios
- ✅ Genera preview en grilla con 25 columnas

**Estadísticas mostradas**:
- Total de filas procesadas
- Cantidad de empresas únicas
- Empleados únicos (por CUIL)
- Warnings generados
- Errores detectados

### 5️⃣ Revisión de Empresas (Modal Automático)

Si se detectan empresas incompletas o múltiples coincidencias:

**Opciones disponibles**:
- 🔍 **Buscar en Empresas.xlsx** - Encuentra por CUIT o nombre
- ✏️ **Editar/Crear empresa** - Completa datos manualmente
- 🗑️ **Eliminar empresa** - Elimina con backup automático

**Búsqueda rápida**:
- Filtrar por CUIT, nombre, localidad o provincia
- Doble click para seleccionar

### 6️⃣ Exportar Resultado

1. Click en **"Exportar XLS"** (se habilita si no hay errores)
2. Elegir ubicación y nombre para el archivo
3. El sistema genera archivo `.xls` con:
   - **Hoja1**: Datos (columnas A-Y)
   - **Hoja2**: Vacía
   - **Hoja3**: Vacía

**Reglas de exportación**:
- Código Postal: siempre vacío (columna E)
- Nro Documento: siempre vacío (columna O)
- Referente: siempre vacío (columna W)

---

## 🏗️ Arquitectura

### Estructura del Proyecto

```
ConvertidorDeOrdenes/
├── 📁 ConvertidorDeOrdenes.Core/      # Capa de lógica (sin UI)
│   ├── 📁 Models/
│   │   ├── OutputRow.cs               # Modelo de salida (A-Y)
│   │   ├── CompanyRecord.cs           # Modelo de empresa
│   │   ├── ParseResult.cs             # Resultado de parseo
│   │   └── ValidationResult.cs        # Resultado de validación
│   ├── 📁 Parsers/
│   │   ├── CsvOrderParser.cs          # Parser CSV
│   │   └── XlsxOrderParser.cs         # Parser XLSX multihoja
│   └── 📁 Services/
│       ├── CompanyRepositoryExcel.cs  # CRUD Empresas.xlsx
│       ├── Normalizer.cs              # Normalización
│       ├── Validator.cs               # Validación
│       ├── PrestacionMapper.cs        # Mapeo prestaciones
│       ├── XlsExporter.cs             # Export NPOI/HSSF
│       └── Logger.cs                  # Sistema de logs
│
├── 📁 ConvertidorDeOrdenes.Desktop/   # Capa de presentación
│   ├── 📁 Forms/
│   │   ├── ArtSelectionForm.cs        # Selección ART
│   │   ├── WizardForm.cs              # Configuración inicial
│   │   ├── MainForm.cs                # Formulario principal
│   │   ├── CompanyResolutionForm.cs   # Revisión de empresas
│   │   ├── CompanyEditDialog.cs       # Alta/edición empresa
│   │   ├── CompanySelectDialog.cs     # Selección múltiple
│   │   └── CompanyListForm.cs         # Administración CRUD
│   └── Program.cs                     # Entry point
│
├── 📁 ParserTester/                   # Consola para probar rápidamente los parsers
│   └── Program.cs                     # Ejemplo de uso de CsvOrderParser
│
├── 📁 logs/                           # Logs generados
├── Empresas.xlsx                      # Base de datos empresas
├── PrestacionesMap.csv                # Mapeo prestaciones
└── README.md                          # Esta documentación
```

### Tecnologías Utilizadas

| Componente | Tecnología | Versión |
|-----------|------------|---------|
| **Framework** | .NET | 8.0 |
| **UI** | Windows Forms | .NET 8 |
| **Excel (lectura)** | ClosedXML | 0.104+ |
| **Excel (escritura XLS)** | NPOI HSSF | 2.7+ |
| **Arquitectura** | Capas (Core + Desktop) | - |

---

## 📊 Modelo de Datos

### Columnas de Salida (A-Y)

| Col | Campo | Req | Descripción | Transformaciones |
|-----|-------|-----|-------------|------------------|
| **A** | CuitEmpleador | ✅ | CUIT del empleador | Validación de longitud (11 dígitos numéricos) |
| **B** | CIIU | ❌ | Código CIIU de actividad | - |
| **C** | Empleador | ✅ | Razón social | Limpieza de formato "NRO - NOMBRE" |
| **D** | Calle | ❌ | Domicilio | - |
| **E** | CodPostal | ❌ | Código postal | ⚠️ Siempre vacío en salida |
| **F** | Localidad | ✅ | Localidad normalizada | Limpieza CP, sufijos provincia |
| **G** | Provincia | ✅ | Provincia normalizada | BA→BUENOS AIRES, CF→CAPITAL FEDERAL |
| **H** | ABMlocProv | ❌ | Alta/Baja/Modif localidad | - |
| **I** | Telefono | ❌ | Teléfono de contacto | - |
| **J** | Fax | ❌ | Número de fax | - |
| **K** | Contrato | ❌ | Número de contrato | - |
| **L** | NroEstablecimiento | ❌ | N° de establecimiento | Extracción desde "NRO - NOMBRE" |
| **M** | Frecuencia | ✅ | A/S/R (Anual/Semestral/Reconf) | Del wizard |
| **N** | Cuil | ✅ | CUIL del trabajador | Validación de longitud (11 dígitos numéricos) |
| **O** | NroDocumento | ❌ | Número de documento | ⚠️ Siempre vacío en salida |
| **P** | TrabajadorApellidoNombre | ✅ | Apellido y nombre completo | - |
| **Q** | Riesgo | ✅ | Descripción del riesgo | Max 90 chars (trunca con warning) |
| **R** | DescripcionRiesgo | ❌ | Descripción extendida | - |
| **S** | ABMRiesgo | ❌ | Alta/Baja/Modif riesgo | - |
| **T** | Prestacion | ✅ | Prestación médica | Limpieza cod:, acentos, mapeo |
| **U** | HistoriaClinica | ❌ | Número de HC | - |
| **V** | Mail | ❌ | Email de contacto | - |
| **W** | Referente | ❌ | Referente | ⚠️ Hoy se exporta vacío, reservada para futuras integraciones |
| **X** | DescripcionError | ❌ | Mensajes de validación | Solo internos |
| **Y** | Id | ❌ | Identificador único | Por definir |

### Leyenda
- ✅ Campo obligatorio (el sistema valida antes de exportar)
- ❌ Campo opcional
- ⚠️ Campo con regla especial de negocio

---

## 🔄 Transformaciones Automáticas

### 1. Prestaciones

**Limpieza de códigos:**
```
Entrada:  "RX DE TORAX DE FRENTE cod: R02"
Salida:   "RX DE TORAX DE FRENTE"
```

**Eliminación de acentos:**
```
"AUDIOMETRÍA TONAL" → "AUDIOMETRIA TONAL"
"ESPIROMETRÍA" → "ESPIROMETRIA"
```

**Mapeo desde PrestacionesMap:**
```csv
# Archivo: PrestacionesMap.csv
ACIDO T-T-MUCONICO EN ORINA,ACIDO TT MUCONICO EN ORINA

# Resultado:
"ACIDO T-T-MUCONICO EN ORINA" → "ACIDO TT MUCONICO EN ORINA"
```

### 2. Empleador y Establecimiento

**Separación de número y nombre:**
```
Entrada:  "2 - Pu innovations srl"
Salida:   NroEstablecimiento = "2"
          Empleador = "Pu innovations srl"
```

**Casos especiales:**
```
"ARCE SEGURIDAD E HIGIENE"  → NroEstablecimiento = ""
                                Empleador = "ARCE SEGURIDAD E HIGIENE"
```

### 3. Localidad

**Extracción de código postal:**
```
"(1605) MUNRO-B A"  → CodPostal = "1605"
                       Localidad = "MUNRO"
```

**Limpieza de sufijos de provincia:**
```
"JUAN BAUTISTA ALBERDI BA"  → "JUAN BAUTISTA ALBERDI"
"SAN MIGUEL CF"             → "SAN MIGUEL"
```

**Casos complejos:**
```
"(6034) LOCALIDAD-B A"  → CodPostal = "6034"
                           Localidad = "LOCALIDAD"
```

### 4. Provincia

**Normalización de abreviaturas:**

| Entrada | Salida |
|---------|--------|
| BA, B A, BS AS, BS. AS. | BUENOS AIRES |
| CF, C.F., CABA, CDAD. DE BS AS, CIUDAD DE BUENOS AIRES | CAPITAL FEDERAL |
| CBA | CORDOBA |
| STA FE, SF | SANTA FE |
| MZA | MENDOZA |
| TUC | TUCUMAN |
| SDE, STGO DEL ESTERO | SANTIAGO DEL ESTERO |
| SL | SAN LUIS |
| SJ | SAN JUAN |

**Limpieza de formato:**
```
"BUENOS AIRES (BA)"  → "BUENOS AIRES"
"Bs. As."            → "BUENOS AIRES"
```

### 5. Riesgo (Frecuencia R)

**Regla especial para Reconfirmatorios:**
```
Si Frecuencia = "R" entonces:
    Riesgo = Prestacion
```

**Truncado con warning:**
```
Si len(Riesgo) > 90:
    Riesgo = Riesgo[0:90]
    Warning: "Riesgo truncado a 90 caracteres para fila X"
```

### 6. CUIL/CUIT

**Validación básica:**
```
"2025913386"        → Warning: "Formato de CUIT posiblemente inválido: 2025913386" (menos de 11 dígitos)
"20-25913386-7"     → OK (11 dígitos numéricos)
```

El sistema conserva el formato de CUIT/CUIL tal como viene en el archivo de entrada; solo verifica que contenga exactamente 11 dígitos numéricos y, si no, genera un warning.

---

## 📝 Sistema de Logs

### Ubicación

```
ConvertidorDeOrdenes/
└── logs/
    ├── log_20260131_093025.txt
    ├── log_20260131_103512.txt
    └── log_20260131_142205.txt
```

Formato de nombre: `log_yyyyMMdd_HHmmss.txt`

### Contenido del Log

```
[09:30:25] [INFO] === Inicio de sesión ===
[09:30:25] [INFO] Tipo de carga: AnualesSemestrales
[09:30:25] [INFO] Frecuencia: A
[09:30:25] [INFO] ART: La Segunda
[09:30:25] [INFO] Referente: 
[09:30:25] [INFO] Empresas.xlsx: C:\Users\...\Empresas.xlsx
[09:30:25] [INFO] Empresas cargadas: 47
[09:30:26] [INFO] Analizando archivo: C:\Users\...\solicitudes_pendiente_prestador9671920260102001935.xlsx
[09:30:27] [INFO] Filas parseadas: 14
[09:30:27] [WARNING] Prestación sin mapeo: RX DE TORAX DE FRENTE
[09:30:27] [WARNING] Riesgo truncado a 90 caracteres para fila 5
[09:30:28] [ERROR] CUIT Empleador es obligatorio (fila 8)
[09:30:32] [INFO] Archivo exportado: C:\Users\...\SALIDA_20260131_093032.xls
```

### Tipos de Mensajes

| Tipo | Prefijo | Descripción |
|------|---------|-------------|
| **INFO** | `[HH:mm:ss] [INFO]` | Operaciones normales |
| **WARNING** | `[HH:mm:ss] [WARNING]` | Advertencias (no bloquean) |
| **ERROR** | `[HH:mm:ss] [ERROR]` | Errores críticos |

---

## 🔧 Solución de Problemas

### ❌ Error: "No se puede abrir Empresas.xlsx"

**Causa**: El archivo está abierto en Excel u otra aplicación

**Solución**:
1. Cerrar Microsoft Excel completamente
2. Verificar que no haya procesos de Excel en el Administrador de Tareas
3. Reintentar la operación

---

### ❌ Error: "Error leyendo archivo CSV"

**Causas posibles**:
- Encoding incorrecto
- Delimitador inválido
- Archivo corrupto

**Soluciones**:
1. Verificar que el CSV use **encoding ISO-8859-1** (Latin-1)
2. Confirmar que el delimitador sea **coma** (,)
3. Abrir en Excel y guardar como "CSV (delimitado por comas)"
4. Verificar que no haya saltos de línea dentro de celdas

---

### ❌ Error: "CUIT Empleador es obligatorio"

**Causa**: No se pudo resolver el CUIT ni desde el archivo ni desde Empresas.xlsx

**Solución**:
1. El sistema abrirá automáticamente el **modal de revisión**
2. Opciones disponibles:
   - 🔍 **Buscar en Empresas.xlsx**: Localizar por nombre
   - ✏️ **Editar/Crear empresa**: Completar CUIT manualmente
   - Si es empresa nueva, ingresar todos los datos

---

### ⚠️ Warning: "Prestación sin mapeo"

**Causa**: La prestación no existe en PrestacionesMap.csv

**Impacto**: No bloquea el proceso, sale sin modificar

**Solución (opcional)**:
1. Agregar entrada a `PrestacionesMap.csv`:
```csv
PRESTACION ORIGINAL,PRESTACION NORMALIZADA
```
2. Procesar nuevamente el archivo

---

### ⚠️ Warning: "Riesgo truncado a 90 caracteres"

**Causa**: El campo Riesgo excede los 90 caracteres permitidos

**Impacto**: Se trunca automáticamente

**Solución**:
- Revisar en el modal de revisión si es necesario
- El sistema guarda log de qué filas fueron truncadas

---

### ❌ Error: "No se detectaron columnas válidas"

**Causa**: El archivo XLSX no tiene el formato esperado

**Soluciones**:
1. Verificar que sea un archivo de La Segunda ART
2. Confirmar que tenga hojas con nombres de solapas
3. Verificar que tenga columna "CUIL" o "Beneficiario"

---

### 🐛 Error: "The process cannot access the file because it is being used"

**Causa**: El proceso previo de la aplicación no cerró correctamente

**Solución**:
1. Abrir Administrador de Tareas (Ctrl+Shift+Esc)
2. Buscar procesos `ConvertidorDeOrdenes.Desktop`
3. Finalizar todos los procesos
4. Compilar/ejecutar nuevamente

---

## 💡 Casos de Uso Comunes

### Caso 1: Procesar Archivo Anual de La Segunda

```
1. Iniciar aplicación
2. Seleccionar ART: "La Segunda"
3. Tipo de carga: "Anuales/Semestrales"
4. Frecuencia: "A - Anual"
5. Elegir archivo XLSX de La Segunda
6. Analizar → Revisar empresas si es necesario
7. Exportar XLS
```

### Caso 2: Agregar Nueva Empresa a la Base

```
Opción A - Durante procesamiento:
1. Al analizar, si no reconoce la empresa → modal de revisión
2. Click "Editar/Crear empresa"
3. Completar datos (CUIT obligatorio)
4. Guardar → queda en Empresas.xlsx

Opción B - Desde menú:
1. Menú: Empresas → Administrar...
2. Click "Agregar"
3. Completar formulario
4. Guardar
```

### Caso 3: Buscar y Eliminar Empresa Duplicada

```
1. Menú: Empresas → Administrar...
2. Usar búsqueda: "nombre empresa"
3. Seleccionar empresa duplicada
4. Click "Eliminar"
5. Confirmar → se crea backup automático
```

---

## 🔐 Seguridad y Backups

### Backups Automáticos

El sistema crea backups automáticamente antes de operaciones destructivas:

**Formato de backup**:
```
Empresas_backup_20260131_142530.xlsx
```

**Cuándo se crea**:
- ✅ Al eliminar una empresa desde el administrador
- ✅ Al eliminar desde el modal de selección múltiple

**Ubicación**:
- Misma carpeta que `Empresas.xlsx`

---

## 📚 Ejemplos de Archivos

> Todos los ejemplos a continuación usan datos completamente ficticios.

### CSV Reconfirmatorios (ejemplo completo)

```csv
Contrato,CUIT,Razón social,Nro. Establecimiento,Nombre establecimiento,Teléfono,CUIL,Nombre Beneficiario,Práctica solicitada,Comentarios,Localidad,Provincia,Teléfono / Celular,Email Beneficiario,Nro Agencia,Email Agencia,Teléfono Agencia
100001,30-00000000-0,"EMPRESA DEMO S.R.L.",1,"Planta Central Demo","1100000000",20-00000000-0,"APELLIDO NOMBRE","ACIDO T-T-MUCONICO EN ORINA cod: L31","Observación de ejemplo","(1000) CIUDAD DEMO","PROVINCIA X","1100000000","empleado@demo-local.test","9999","agencia@demo-local.test","1100000000"
100002,30-00000000-0,"EMPRESA DEMO S.R.L.",1,"Planta Central Demo","1100000000",27-00000000-0,"OTRO APELLIDO","EXAMEN CLINICO PREOCUPACIONAL cod: C06","","(1000) CIUDAD DEMO","PROVINCIA X","1100000000","empleado2@demo-local.test","9999","agencia@demo-local.test","1100000000"
```

### XLSX Anuales/Semestrales (estructura)

**Hoja "Resumen"** (ejemplo):
| Solapa | Razón Social        | CUIT          |
|--------|---------------------|---------------|
| 1      | EMPRESA DEMO S.R.L. | 30-00000000-0 |
| 2      | EMPRESA PRUEBA SA   | 30-11111111-1 |

**Hoja "1"** (datos empleados ficticios):
| CUIL         | Beneficiario      | Riesgo         | Examen                 |
|--------------|-------------------|----------------|------------------------|
| 20-00000000-0| APELLIDO NOMBRE   | ADMINISTRATIVO | C06 - EXAMEN CLINICO   |
| 27-00000000-0| OTRO APELLIDO     | OPERARIO       | C06 - EXAMEN CLINICO   |

---

## 🚀 Roadmap Futuro

- [ ] Soporte para múltiples ARTs (Pepito, etc.)
- [ ] Columna ID con lógica de autoincremento
- [ ] Scraping de datos desde portales de ART
- [ ] Sistema de auto-actualización
- [ ] Versión web progresiva
- [ ] Exportación a formatos adicionales (XLSX, PDF)
- [ ] Importación masiva de empresas desde Excel
- [ ] Dashboard de estadísticas
- [ ] Integración con bases de datos SQL

---

## 👥 Soporte

Para reportar problemas o solicitar nuevas características:

1. **Logs**: Revisar carpeta `logs/` y adjuntar el archivo más reciente
2. **Datos**: Incluir ejemplo de archivo de entrada (sin datos sensibles)
3. **Pasos**: Describir paso a paso para reproducir el problema

---

## 📄 Licencia

**Uso Interno - Seres Salud**

Este software es propiedad de Seres Salud y está destinado exclusivamente para uso interno. Queda prohibida su distribución, modificación o uso comercial sin autorización expresa.

---

## 🏆 Créditos

**Desarrollado para**: Seres Salud  
**Framework**: .NET 8  
**UI**: Windows Forms  
**Excel**: ClosedXML + NPOI  

---

<div align="center">

**ConvertidorDeOrdenes v2.0**  
© 2026 Seres Salud - Todos los derechos reservados

</div>

**Versión**: 1.0.0  
**Fecha**: Enero 2026  
**Framework**: .NET 8 + WinForms
# SeresSalud
