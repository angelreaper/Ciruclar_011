# Procesador de Archivos Excel - VBA

## 1. Descripción

Este proyecto consiste en un archivo Excel habilitado para macros (`.xlsm`) que permite seleccionar un archivo Excel de origen, leer y procesar su información, aplicar reglas de validación y homologación, transformar cada registro a una estructura de salida y generar archivos `.txt` separados por fondo.

La herramienta está desarrollada en **VBA (Visual Basic for Applications)** y actualmente corresponde a un prototipo funcional sobre el cual se continuará construyendo la lógica definitiva del proceso.

---

## 2. Objetivo

El objetivo principal es automatizar el procesamiento de información proveniente de archivos Excel.

El flujo general es:

```text
Archivo Excel origen
        │
        ▼
Selección del archivo
        │
        ▼
Validación de estructura
        │
        ▼
Lectura de información
        │
        ▼
Validación y homologación
        │
        ▼
Transformación del registro
        │
        ▼
Generación de registro de salida
        │
        ▼
Generación de contrapartida
        │
        ▼
Archivo TXT por fondo
```

---

# 3. Interfaz del archivo

El libro contiene las siguientes hojas:

| Hoja            | Descripción                                                              |
| --------------- | ------------------------------------------------------------------------ |
| `Inicio`        | Pantalla principal para seleccionar el archivo y ejecutar el proceso.    |
| `Configuracion` | Espacio destinado para parámetros configurables del proceso.             |
| `Homologacion`  | Contiene las equivalencias entre fondo, portafolio y código SIFI.        |
| `Resultado`     | Permite visualizar información relacionada con los registros procesados. |
| `Log`           | Registra eventos, advertencias y errores ocurridos durante la ejecución. |

---

# 4. Estructura VBA

El proyecto se encuentra organizado en módulos según la responsabilidad de cada funcionalidad:

```text
VBAProject
│
├── ModMain
├── ModArchivos
├── ModExcel
├── ModLog
├── ModTxt
├── ModProgreso
├── ModUtilidades
├── ModValidaciones
├── ModHomologacion
├── ModTransformacion
└── FrmProgreso
```

## 4.1 ModMain

Es el módulo principal de ejecución.

Responsabilidades:

* Seleccionar el archivo de origen.
* Obtener la ruta del archivo.
* Abrir el archivo.
* Obtener la hoja a procesar.
* Leer la información.
* Validar la estructura.
* Iniciar el procesamiento.
* Controlar la barra de progreso.
* Controlar la cancelación del proceso.
* Generar los archivos de salida.
* Mostrar mensajes al usuario.
* Abrir la carpeta de resultados al finalizar.

Principales procedimientos:

```vba
SeleccionarArchivoOrigen()
Procesar()
ProcesarRegistro()
```

---

## 4.2 ModArchivos

Responsable de abrir y cerrar los archivos Excel utilizados durante el proceso.

Funciones principales:

```vba
AbrirArchivo()
CerrarArchivo()
```

El archivo de origen se abre en modo:

```text
ReadOnly = True
```

para evitar modificaciones sobre el archivo original.

---

## 4.3 ModExcel

Responsable de la lectura y manipulación básica de información de Excel.

Funciones principales:

```vba
ObtenerHojas()
LeerHoja()
ObtenerUltimaFila()
ObtenerUltimaColumna()
LimpiarResultado()
ConstruirTextoRegistro()
```

La información de la hoja se carga en memoria mediante un `Variant`, permitiendo trabajar con los registros sin realizar una lectura celda por celda durante todo el procesamiento.

---

## 4.4 ModValidaciones

Responsable de validar que el archivo de origen tenga las columnas necesarias.

Actualmente se consideran como columnas requeridas:

```text
FONDO
ENCARGO
OPIN
TIPO_MOV
TIPO_MOV_DESC
FECHA_MOV
VALOR
```

Funciones principales:

```vba
ValidarEstructura()
ObtenerColumna()
```

`ObtenerColumna()` permite localizar una columna por su nombre, independientemente de la posición que tenga dentro del archivo de origen.

Por ejemplo, si `FONDO` está en la columna A o en la columna H, el proceso puede localizarla por nombre.

---

# 5. Estructura esperada del archivo de origen

El archivo de origen contiene información como:

```text
FONDO
FONDO_DESC
ENCARGO
ENCARGO_DESC
EMPRESA
OPIN
OPIN_DESC
NUMERO_MOVI
TIPO_MOV
TIPO_MOV_DESC
OFICINA
FECHA_MOV
MOV_EFECTIVO
VALOR
```

No todas las columnas son utilizadas actualmente para la transformación.

Las principales columnas utilizadas son:

```text
FONDO
ENCARGO
OPIN
TIPO_MOV
TIPO_MOV_DESC
FECHA_MOV
VALOR
```

---

# 6. Homologación de fondos

La hoja `Homologacion` contiene la relación entre el fondo de origen, el portafolio y el código SIFI.

Estructura:

| Fondo        | Portafolio |   SIFI |
| ------------ | ---------: | -----: |
| Sumar        |         60 |  22969 |
| Fidugob      |       1090 |  10659 |
| Impulso      |        580 | 125575 |
| ES+          |        191 |  54077 |
| Altarenta    |        404 |  70843 |
| Cubrir       |         31 |  29497 |
| Óptimo       |         19 |  29495 |
| ETF          |        380 |  43502 |
| ID ETF       |        N/A | 118898 |
| Rentar       |        700 |  11286 |
| Fidulíquidez |        701 |  11356 |
| Rentar 30    |        702 | 129224 |

La homologación permite convertir el valor de `FONDO` del archivo de origen en el código SIFI requerido para el archivo de salida.

Ejemplo:

```text
FONDO origen
60
 │
 ▼
Portafolio
Sumar
 │
 ▼
SIFI
22969
```

Si un fondo no existe en la tabla de homologación, el registro se considera una excepción y se registra en el `Log`.

---

# 7. Transformación del registro

La transformación se encuentra principalmente en:

```text
ModTransformacion
```

La función principal es:

```vba
TransformarRegistro()
```

La salida está compuesta por **17 columnas**, separadas mediante `|`.

La estructura es:

| Columna | Origen / Regla                                  |
| ------: | ----------------------------------------------- |
|       1 | SIFI homologado a partir de `FONDO`             |
|       2 | `FECHA_MOV` → `AAAAMM`                          |
|       3 | Valor fijo `14`                                 |
|       4 | `FECHA_MOV` → `AAAAMMDD`                        |
|       5 | `TIPO_MOV_DESC + TIPO_MOV`                      |
|       6 | `C` si el valor es positivo, `D` si es negativo |
|       7 | Valor fijo `3525100101`                         |
|       8 | Valor absoluto de `VALOR`                       |
|       9 | `ENCARGO` eliminando comas                      |
|      10 | Vacío                                           |
|      11 | Año de `FECHA_MOV`                              |
|      12 | Vacío                                           |
|      13 | Vacío                                           |
|      14 | Valor fijo `MO`                                 |
|      15 | Vacío                                           |
|      16 | `OPIN`                                          |
|      17 | `OPIN`                                          |

---

# 8. Ejemplo de transformación

## Datos de entrada

```text
FONDO       = 60
ENCARGO     = 6000200100104
OPIN        = 6014
TIPO_MOV    = 1
TIPO_MOV_DESC = APORTES INVERSIONISTAS
FECHA_MOV   = 21/07/2026
VALOR       = 1000000
```

El fondo `60` se homologa a:

```text
SIFI = 22969
```

El registro resultante es:

```text
22969|202607|14|20260721|APORTES INVERSIONISTAS 1|C|3525100101|1000000|6000200100104||2026|||MO||6014|6014
```

---

# 9. Generación de contrapartida

Cada registro procesado genera dos líneas de salida.

La segunda línea corresponde a la contrapartida del movimiento.

La regla es:

```text
C → D
D → C
```

Todos los demás campos permanecen iguales.

## Ejemplo

Registro original:

```text
22969|202607|14|20260721|APORTES INVERSIONISTAS 1|C|3525100101|1000000|6000200100104||2026|||MO||6014|6014
```

Contrapartida:

```text
22969|202607|14|20260721|APORTES INVERSIONISTAS 1|D|3525100101|1000000|6000200100104||2026|||MO||6014|6014
```

Por lo tanto, un registro de entrada genera:

```text
Registro original
        +
Contrapartida
```

Las dos líneas se escriben consecutivamente en el archivo de salida.

---

# 10. Regla de naturaleza del movimiento

La naturaleza del movimiento se determina a partir del campo `VALOR`.

```text
VALOR > 0
   ↓
   C
   ↓
Contrapartida D
```

Mientras que:

```text
VALOR < 0
   ↓
   D
   ↓
Contrapartida C
```

El valor que se escribe en la columna 8 siempre es absoluto.

Por ejemplo:

```text
VALOR = -1000000
```

genera:

```text
D|3525100101|1000000
```

y su contrapartida:

```text
C|3525100101|1000000
```

---

# 11. Archivos de salida

Los resultados se generan dentro de la carpeta `Downloads` del usuario.

La carpeta se crea con una estructura similar a:

```text
Downloads
└── RESULTADO_PROCESO_27082026
    ├── FONDO_SUMAR_270826_213500.txt
    ├── FONDO_FIDUGOB_270826_213500.txt
    └── FONDO_IMPULSO_270826_213500.txt
```

Cada fondo tiene su propio archivo de salida.

Los archivos generados:

* No contienen encabezados.
* Contienen una línea por registro generado.
* Incluyen la línea original y su contrapartida.
* Utilizan `|` como separador.

---

# 12. ModTxt

Responsable de la generación de archivos de texto.

Funciones principales:

```vba
ObtenerRutaResultados()
CrearCarpetaResultados()
ObtenerRutaTxt()
EscribirLineaTxt()
```

La escritura se realiza utilizando modo `Append`, permitiendo agregar consecutivamente nuevos registros al archivo correspondiente.

---

# 13. Barra de progreso

El formulario:

```text
FrmProgreso
```

permite visualizar el avance del procesamiento.

Muestra:

```text
Preparando...

45 %

Procesando registros...

450 de 1000

[████████████░░░░░░░░]
```

El usuario también puede cancelar el procesamiento.

La funcionalidad está implementada mediante:

```vba
MostrarProgreso()
ActualizarProgreso()
ProcesoCancelado()
CerrarProgreso()
```

El uso de `DoEvents` permite que Excel pueda responder a las acciones del usuario mientras se ejecuta el procesamiento.

---

# 14. Cancelación del proceso

Durante el procesamiento se verifica periódicamente si el usuario ha solicitado cancelar.

Conceptualmente:

```text
Procesar registro
      ↓
¿Cancelado?
  ├── Sí → detener proceso
  └── No → continuar
```

Cuando el usuario cancela:

* Se registra el evento en `Log`.
* Se detiene el procesamiento.
* Se cierra el formulario de progreso.
* Se cierra el archivo origen.
* Se restauran las opciones de Excel.

---

# 15. Log de procesamiento

La hoja:

```text
Log
```

permite registrar los eventos ocurridos durante la ejecución.

Los tipos de eventos utilizados son:

```text
INFO
OK
ADVERTENCIA
ERROR
```

Ejemplos:

```text
INFO        Inicio del procesamiento
INFO        Archivo abierto: ArchivoOrigen.xlsx
INFO        Procesando hoja: Hoja1
INFO        Registros encontrados: 1500
ADVERTENCIA Proceso cancelado por el usuario
ERROR       Fondo no homologado: 999
ERROR       Fecha inválida
```

El log permite identificar problemas sin interrumpir innecesariamente todo el procesamiento.

---

# 16. Apertura automática de resultados

Al finalizar correctamente el proceso, la herramienta muestra un mensaje indicando la ubicación de los resultados.

Posteriormente abre automáticamente la carpeta:

```text
RESULTADO_PROCESO_ddmmyyyy
```

en el Explorador de Windows.

Esto permite al usuario acceder inmediatamente a los archivos generados.

---

# 17. Manejo de errores

El procedimiento principal cuenta con un manejador de errores.

Ante una excepción:

1. Se registra el número y descripción del error.
2. Se intenta cerrar el formulario de progreso.
3. Se intenta cerrar el archivo de origen.
4. Se restauran las opciones de Excel.
5. Se muestra un mensaje al usuario.

Ejemplo:

```text
Se produjo un error:

1004 - Descripción del error
```

---

# 18. Consideraciones actuales

Esta implementación corresponde a un **prototipo funcional**.

Actualmente se encuentra definido:

* Selección del archivo.
* Apertura del archivo.
* Lectura de información.
* Validación básica de estructura.
* Homologación de fondos.
* Transformación de registros.
* Generación de contrapartidas.
* Generación de archivos TXT.
* Separación de archivos por fondo.
* Barra de progreso.
* Cancelación.
* Log.
* Apertura automática de la carpeta de resultados.

Algunas reglas funcionales pueden requerir ajustes cuando se conozca completamente la estructura definitiva de los archivos de origen y destino.

---

# 19. Flujo completo

El flujo actual puede resumirse de la siguiente manera:

```text
┌──────────────────────┐
│     Usuario          │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Seleccionar archivo  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Abrir archivo Excel  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Leer primera hoja    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Validar estructura   │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Obtener columnas     │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Procesar registros   │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Homologar FONDO      │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Transformar registro │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Generar C / D        │
│ y contrapartida      │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Escribir TXT         │
│ por fondo             │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Registrar Log        │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Finalizar proceso    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Abrir carpeta        │
│ de resultados        │
└──────────────────────┘
```

---

# 20. Próximas mejoras

Una vez definida completamente la estructura de los archivos de entrada y las reglas de negocio, se pueden implementar mejoras como:

* Validaciones adicionales de datos.
* Identificación de registros duplicados.
* Control detallado de registros exitosos y fallidos.
* Resumen de procesamiento por fondo.
* Contadores de registros procesados, exitosos y con error.
* Mayor control sobre archivos de salida.
* Optimización para archivos con grandes volúmenes de registros.
* Validación de tipos de datos.
* Manejo de múltiples hojas de origen.
* Configuración de parámetros desde la hoja `Configuracion`.
* Mejoras en el manejo de errores.
* Generación de un archivo consolidado de excepciones.
* Protección del proyecto VBA para la versión de distribución.

---

## 21. Versión

**Versión actual:** `1.0`

**Estado:** Prototipo funcional en desarrollo.
