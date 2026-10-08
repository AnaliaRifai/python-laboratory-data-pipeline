# 🧪 Pipeline en Python para normalización de informes históricos de laboratorio

> ETL con Python · Limpieza de datos históricos · Normalización de archivos Excel · Validación de calidad · Consolidación 2018–2026

---

## 🛠️ Stack tecnológico

| Herramienta / Librería | Uso |
|---|---|
| Python | Desarrollo del pipeline de procesamiento |
| Pandas | Lectura, transformación, validación y consolidación de datos |
| OpenPyXL | Lectura y escritura de archivos `.xlsx` |
| xlrd | Lectura de archivos históricos `.xls` |
| lxml | Lectura alternativa de archivos `.xls` exportados como HTML |
| re | Normalización de textos y encabezados |
| os / glob | Búsqueda automática de archivos en carpetas |

---

## 🎯 Objetivo del proyecto

El objetivo general del proyecto es construir una base histórica confiable de resultados de laboratorio que permita consultar y analizar información acumulada durante varios años.

Originalmente, los informes se encontraban distribuidos en múltiples archivos Excel con estructuras variables, lo que dificultaba responder preguntas simples como:

- ¿Qué muestras se analizaron en determinado año?
- ¿Qué resultados tuvo un productor específico?
- ¿Qué cultivos o alimentos fueron analizados con mayor frecuencia?
- ¿Cómo consultar rápidamente resultados históricos sin revisar archivo por archivo?

Este repositorio documenta la **primera etapa del proyecto**: el desarrollo de un pipeline en Python para normalizar, validar y consolidar los informes históricos del período 2018–2026.

El resultado de esta etapa es una base histórica limpia y estandarizada, que funciona como insumo para una etapa posterior de análisis, consulta y visualización.

---

## 📌 Contexto

El laboratorio contaba con informes históricos generados en Excel durante varios años. Estos archivos no tenían una estructura completamente uniforme, lo que dificultaba su consolidación directa.

Entre los principales problemas detectados se encontraban:

- Archivos en formatos `.xlsx` y `.xls`.
- Encabezados ubicados en distintas filas.
- Encabezados de una o dos filas.
- Diferentes nombres para una misma variable analítica.
- Metadatos ubicados antes de la tabla principal.
- Columnas fuera del alcance del proyecto.
- Valores faltantes representados de distintas formas.
- Informes con tablas ubicadas horizontalmente.
- Archivos acumulativos o versiones incompletas.
- Errores de escritura y abreviaturas históricas.

Por este motivo, se decidió utilizar Python como herramienta principal para construir una etapa de normalización previa al análisis.

---

## 🧭 Alcance de esta etapa

Este repositorio corresponde a la etapa de preparación y consolidación de datos.

El foco de esta etapa fue resolver el problema de base: transformar archivos históricos dispersos y heterogéneos en una estructura única, confiable y reutilizable.

Incluye:

- lectura automática de informes históricos;
- normalización de estructuras diferentes;
- estandarización de nombres de variables;
- tratamiento de errores, valores faltantes y excepciones;
- validación anual de los datos;
- consolidación final del período 2018–2026;
- generación de logs de control.

No incluye todavía la etapa completa de visualización o dashboard.  
Esa etapa se desarrollará a partir de la base histórica generada por este pipeline.

---

## 🗂️ Estructura del repositorio

```text
python-laboratory-data-pipeline/
│
├── scripts/
│   ├── pipeline_2018.py
│   └── pipeline_maestro_consolidado_historico_2018_2026.py
│
├── docs/
│   └── bitacora_tecnica.md
│
├── images/
│
├── outputs_sample/
│
└── README.md
```

### Descripción de carpetas

| Carpeta | Contenido |
|---|---|
| `scripts/` | Scripts principales del proyecto |
| `docs/` | Documentación técnica complementaria |
| `images/` | Capturas utilizadas en la documentación |
| `outputs_sample/` | Ejemplos anonimizados o capturas de salidas generadas |

---

## 🔄 Flujo general de esta etapa

```text
Archivos Excel originales
        ↓
Pipeline anual en Python
        ↓
Consolidado anual validado
        ↓
Script maestro histórico
        ↓
Base histórica 2018–2026
        ↓
Dataset preparado para consulta y análisis
```

Los archivos originales se conservaron sin modificaciones.

Todas las reglas de limpieza, corrección y exclusión quedaron documentadas dentro del pipeline.

La base resultante será utilizada en una etapa posterior para construir una herramienta de consulta y visualización que permita explorar la información histórica de forma más rápida y accesible.

---

## 🧩 Desarrollo del pipeline

### Caso base: año 2018

El año 2018 fue utilizado como caso base para construir la lógica principal del pipeline, ya que presentaba varias diferencias estructurales entre archivos.

Durante esta etapa se desarrollaron funciones para:

- Leer archivos `.xlsx` y `.xls`.
- Detectar automáticamente la fila donde comienza la tabla de muestras.
- Identificar si el encabezado ocupa una o dos filas.
- Normalizar nombres de columnas.
- Extraer metadatos del informe.
- Filtrar únicamente filas correspondientes a muestras reales.
- Convertir variables analíticas a formato numérico.
- Excluir columnas fuera del alcance del proyecto.
- Registrar avisos, errores y controles de trazabilidad.
- Exportar un consolidado anual y un log de control.

---

## 🔎 Funciones principales del pipeline

### `limpiar_texto()`

Normaliza los encabezados para poder compararlos aunque estén escritos de forma distinta.

Ejemplos:

```text
% Almidón   → almidon
Muestra Nº  → muestrano
ENERGIA Mcal/kg MS → energiamcalkgms
```

Esta función permite luego buscar cada encabezado dentro de un diccionario de equivalencias.

### `buscar_encabezado()`

Recorre el archivo fila por fila y columna por columna hasta encontrar dónde comienza la tabla de muestras.

Esto permite procesar archivos donde la tabla no siempre empieza en la misma fila.

### `extraer_metadatos()`

Extrae información general ubicada antes de la tabla principal, como:

- solicitante;
- localidad;
- fecha de ingreso;
- fecha de egreso.

Estos datos se agregan luego a cada muestra procesada.

### `procesar_reporte()`

Es la función principal del pipeline anual.

Su objetivo es procesar un informe individual y devolver una tabla estandarizada con sus muestras.

Entre sus tareas se incluyen:

- leer el archivo;
- ubicar el encabezado;
- construir nombres de columnas estandarizados;
- eliminar filas que no corresponden a muestras;
- convertir resultados analíticos a formato numérico;
- agregar metadatos;
- registrar avisos;
- devolver un DataFrame limpio.

---

## 🧱 Normalización de columnas

Se utilizó un diccionario llamado `MAPEO_COLUMNAS` para unificar diferentes nombres históricos que representaban una misma variable.

| Nombre detectado | Nombre estandarizado |
|---|---|
| `LDAs` | `lignina_pct` |
| `% NT` | `nt_pct` |
| `% PB` | `pb_pct` |
| `% EE` | `extracto_etereo_pct` |
| `% N` | `nt_pct` |
| `ENERGIA Mcal/kg MS` | `em_mcal_kg_ms` |

Esto permitió consolidar información histórica aunque los encabezados no fueran idénticos entre años.

---

## 🧪 Problemas de calidad detectados y soluciones aplicadas

| Problema detectado | Solución implementada |
|---|---|
| Encabezados en diferentes filas | Detección automática del inicio de la tabla |
| Encabezados de una o dos filas | Regla dinámica según el contenido debajo de `Muestra` |
| Archivos `.xls` y `.xlsx` | Lectura con librerías específicas según formato |
| Variables con distintos nombres históricos | Diccionario de equivalencias `MAPEO_COLUMNAS` |
| Columnas fuera del alcance | Exclusión documentada mediante `COLUMNAS_EXCLUIDAS` |
| Archivos duplicados o acumulativos | Exclusión controlada mediante `ARCHIVOS_EXCLUIDOS` |
| Valores dobles en una celda | Separación, conversión y cálculo del promedio |
| Guion `-` como ausencia de dato | Conversión a valor faltante (`NaN`) |
| Tablas horizontales dentro de una hoja | Procesamiento por bloques y unificación vertical |
| Metadatos ausentes en archivos originales | Conservación como valores faltantes y registro en log |
| Errores ortográficos o abreviaturas | Incorporación de nuevas reglas de mapeo |

---

## 📆 Adaptación del pipeline por año

Una vez validada la lógica del pipeline con 2018, se replicó y ajustó progresivamente para los años 2019 a 2026.

Cada año fue procesado de forma individual. Luego se revisaron los logs generados para identificar errores, avisos o nuevas estructuras no contempladas.

| Año | Mejora aplicada |
|---|---|
| 2019 | Tratamiento de nuevas variables históricas y estructuras especiales |
| 2020 | Promedio de valores dobles en una misma celda y corrección documentada de metadatos |
| 2021 | Exclusión de variantes de la columna `Peso seco` |
| 2022 | Interpretación de `-` como dato faltante y procesamiento de tres tablas horizontales |
| 2023 | Ejecución sin nuevas incidencias |
| 2024 | Mapeo de `%EE` como `extracto_etereo_pct` |
| 2025 | Ejecución sin nuevas incidencias |
| 2026 | Mapeo de `%N` como `nt_pct` y conservación de metadatos faltantes reales |

---

## 🧾 Script maestro histórico

Luego de validar los consolidados anuales, se desarrolló un script maestro para unificar todos los años en una única base histórica.

El script maestro realiza las siguientes tareas:

- Lee los consolidados anuales de 2018 a 2026.
- Agrega una columna `anio`.
- Une todos los archivos mediante Pandas.
- Estandariza nombres de columnas.
- Renombra `observaciones` como `cultivo`.
- Estandariza `cultivo` en mayúsculas.
- Convierte fechas a formato fecha.
- Redondea variables analíticas a dos decimales.
- Elimina columnas no necesarias para la base final.
- Trata columnas sin nombre.
- Genera controles de calidad históricos.
- Exporta una base final y un log de control.

---

## ✅ Controles de calidad

El pipeline genera controles tanto a nivel anual como histórico.

Entre los controles aplicados se encuentran:

- cantidad de archivos procesados;
- cantidad de muestras consolidadas;
- errores de procesamiento;
- números de muestra duplicados;
- fechas de egreso anteriores a la fecha de ingreso;
- fechas fuera del período esperado;
- valores fuera de rangos analíticos;
- valores nulos por columna;
- cobertura de variables analíticas;
- trazabilidad por archivo;
- avisos y errores en archivos de log.

---

## 📤 Salidas generadas

El proceso genera dos archivos principales:

```text
consolidado_historico_2018_2026.xlsx
log_consolidado_historico_2018_2026.xlsx
```

### Base histórica final

Contiene la información consolidada y estandarizada de los informes históricos del laboratorio.

Incluye variables como:

- `anio`
- `muestra`
- `solicitante`
- `localidad`
- `fecha_ingreso`
- `fecha_egreso`
- `cultivo`
- `ms_pct`
- `nt_pct`
- `pb_pct`
- `fdn_pct`
- `fda_pct`
- `lignina_pct`
- `divms_pct`
- `em_mcal_kg_ms`
- `extracto_etereo_pct`
- `celulosa_pct`

### Log de control

Contiene información de validación del proceso:

- resumen por año;
- columnas por consolidado anual;
- valores nulos por columna;
- cobertura de variables analíticas;
- duplicados;
- fechas inconsistentes;
- columnas sin nombre tratadas;
- avisos;
- errores.

---

## 🔐 Confidencialidad de los datos

Por tratarse de información histórica real de laboratorio, los archivos originales y la base completa no se publican en este repositorio.

El repositorio contiene los scripts principales y documentación del proceso.

Las salidas incluidas, en caso de agregarse, serán ejemplos anonimizados o capturas parciales.

---

## ▶️ Cómo ejecutar el proyecto

Para ejecutar los scripts en otro equipo, se debe adaptar la ruta del proyecto dentro de los archivos `.py`.

Ejemplo:

```python
RAIZ_PROYECTO = r"RUTA_A_TU_PROYECTO\Proyecto Laboratorio"
```

### Librerías necesarias

```bash
pip install pandas openpyxl xlrd lxml
```

### Orden sugerido

1. Ejecutar el pipeline anual base:

```text
scripts/pipeline_2018.py
```

2. Una vez generados y validados los consolidados anuales, ejecutar:

```text
scripts/pipeline_maestro_consolidado_historico_2018_2026.py
```

---

## 📌 Estado del proyecto

Estado actual: **Etapa 1 finalizada**.

Esta etapa incluyó:

- normalización de archivos históricos;
- validación anual del período 2018–2026;
- consolidación histórica final;
- generación de logs de control;
- construcción de una base limpia y reutilizable.

Próxima etapa:

- diseño de una herramienta de consulta y visualización;
- creación de métricas e indicadores;
- desarrollo de filtros por año, productor, cultivo y muestra;
- análisis exploratorio de la información histórica.

---

## 💡 Aprendizajes principales

Este proyecto permitió trabajar con un caso real de datos históricos no estandarizados.

Los principales aprendizajes fueron:

- diseñar un pipeline reproducible en Python;
- separar datos originales de datos procesados;
- documentar reglas de negocio y excepciones;
- construir funciones reutilizables;
- validar resultados mediante logs;
- tratar datos faltantes sin reemplazarlos incorrectamente por cero;
- resolver diferencias históricas de formato y nomenclatura;
- generar una base analítica confiable a partir de archivos operativos desordenados.

---

