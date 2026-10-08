# Bitácora técnica del proyecto

## Pipeline en Python para consolidación histórica de informes de laboratorio

Esta bitácora documenta el proceso técnico desarrollado durante la primera etapa del proyecto: la normalización, validación y consolidación de informes históricos de laboratorio correspondientes al período 2018–2026.

El objetivo de esta etapa fue transformar múltiples archivos Excel con estructuras variables en una base histórica única, limpia y reutilizable para análisis posterior.

---

## 1. Inicio del pipeline en Python

El proyecto comenzó con la necesidad de procesar informes históricos almacenados en archivos Excel. Estos archivos no tenían una estructura completamente homogénea, por lo que no era conveniente consolidarlos manualmente ni cargarlos directamente en una herramienta de visualización.

Se decidió utilizar **Python + Pandas** como etapa previa de procesamiento debido a la variabilidad estructural de los archivos históricos.

La lógica general definida fue:

```text
Archivos Excel originales
        ↓
Python / Pandas
        ↓
Base estandarizada anual
        ↓
Base histórica consolidada
        ↓
Análisis y consulta posterior
```

Desde el inicio se definieron dos criterios importantes:

- los archivos originales debían conservarse sin modificaciones;
- las reglas de limpieza, exclusión y corrección debían quedar documentadas dentro del pipeline.

---

## 2. Organización inicial del proyecto

Se trabajó con una estructura de carpetas separando archivos originales, respaldos, procesados y scripts.

La organización conceptual fue:

```text
Proyecto Laboratorio/
│
├── DATOS/
│   ├── ORIGINALES/
│   │   ├── 2018/
│   │   ├── 2019/
│   │   ├── ...
│   │   └── 2026/
│   │
│   ├── RESPALDOS/
│   │   ├── consolidado_2018.xlsx
│   │   ├── consolidado_2019.xlsx
│   │   └── ...
│   │
│   └── PROCESADOS/
│       ├── consolidado_historico_2018_2026.xlsx
│       └── log_consolidado_historico_2018_2026.xlsx
│
└── PYTHON/
    ├── pipeline_2018.py
    └── pipeline_maestro_consolidado_historico_2018_2026.py
```

Esta separación permitió mantener trazabilidad entre:

- archivos originales;
- resultados anuales;
- base histórica final;
- logs de validación;
- scripts utilizados.

---

## 3. Caso base: procesamiento de 2018

El año 2018 se utilizó como caso base para construir la lógica principal del pipeline.

Durante la exploración inicial se detectó que los archivos presentaban diferencias importantes:

- no todos tenían la misma extensión;
- algunos estaban en formato `.xlsx`;
- otros estaban en formato `.xls`;
- los encabezados no siempre comenzaban en la misma fila;
- algunos encabezados ocupaban una fila y otros dos;
- los metadatos estaban ubicados antes de la tabla principal;
- algunas columnas tenían nombres históricos diferentes para una misma variable.

Por ese motivo, se decidió no asumir posiciones fijas dentro de los archivos.

En lugar de eso, se desarrolló una lógica dinámica para detectar la estructura de cada informe.

---

## 4. Detección de formatos de archivo

Durante la validación de la carpeta 2018 se detectaron archivos en dos formatos principales:

- `.xlsx`
- `.xls`

Inicialmente, el script solo buscaba archivos `.xlsx`, por lo que no encontraba la totalidad de los informes.

Luego se modificó la búsqueda para contemplar ambos formatos:

```python
archivos = (
    glob.glob(os.path.join(carpeta_anio, "*.xlsx")) +
    glob.glob(os.path.join(carpeta_anio, "*.xls"))
)
```

Este punto permitió diferenciar dos etapas del proceso:

1. **Localizar archivos** dentro de una carpeta.
2. **Leer correctamente su contenido** con la librería adecuada.

La detección de archivos con `glob` podía funcionar correctamente, pero para leer archivos `.xls` era necesario contar con dependencias específicas, como `xlrd`.

Este fue uno de los primeros aprendizajes técnicos del proyecto: encontrar archivos y poder leerlos no son procesos equivalentes.

---

## 5. Funciones principales desarrolladas

### 5.1 `limpiar_texto()`

Se creó una función para normalizar los encabezados antes de compararlos con un diccionario de equivalencias.

La función permite reducir diferencias causadas por:

- tildes;
- símbolos;
- espacios;
- mayúsculas y minúsculas;
- signos como `%`, `º`, `/`, entre otros.

Ejemplos:

```text
% Almidón       → almidon
Muestra Nº      → muestrano
ENERGIA Mcal/kg MS → energiamcalkgms
```

Esta normalización permitió comparar encabezados históricos escritos de distintas formas contra claves estándar del pipeline.

---

### 5.2 `buscar_encabezado()`

Esta función recorre filas y columnas del archivo hasta localizar una celda que contenga la palabra `"muestra"`.

Su objetivo es detectar automáticamente dónde comienza la tabla principal del informe.

Esto fue necesario porque la tabla de muestras no siempre comenzaba en la misma fila ni en la misma posición.

La función devuelve:

- la fila donde comienza el encabezado;
- la columna donde comienza la tabla.

Si no encuentra una celda válida, el pipeline registra un error para evitar procesar incorrectamente el archivo.

---

### 5.3 `extraer_metadatos()`

La función `extraer_metadatos()` se creó para recuperar información general ubicada antes de la tabla principal.

Los metadatos buscados fueron:

- solicitante;
- localidad;
- fecha de ingreso;
- fecha de egreso.

La función recorre únicamente las celdas anteriores al encabezado detectado, identifica textos con estructura `Etiqueta: contenido` y guarda esos valores en un diccionario.

Ejemplo conceptual:

```text
Solicitante: XXXX
Localidad: XXXX
Fecha de ingreso: DD/MM/AAAA
Fecha de egreso: DD/MM/AAAA
```

Luego, estos metadatos se agregan como columnas a cada muestra correspondiente al informe procesado.

---

### 5.4 `procesar_reporte()`

La función `procesar_reporte()` se convirtió en la función principal del pipeline anual.

Su objetivo es recibir la ruta de un archivo Excel y devolver un DataFrame limpio y estandarizado.

Las tareas principales de esta función son:

1. Leer el archivo sin asumir encabezados fijos.
2. Detectar la posición del encabezado.
3. Eliminar columnas anteriores a la tabla de muestras.
4. Determinar si el encabezado ocupa una o dos filas.
5. Construir nombres de columnas estandarizados.
6. Extraer únicamente las filas correspondientes a muestras reales.
7. Convertir variables analíticas a formato numérico.
8. Incorporar metadatos del informe.
9. Agregar trazabilidad del archivo original.
10. Registrar avisos y errores.

---

## 6. Detección de encabezados de una o dos filas

Se detectó que algunos informes tenían encabezados de una fila, mientras que otros tenían encabezados de dos filas.

Para resolverlo, se agregó una regla dinámica:

- si el valor ubicado debajo de `"Muestra"` era numérico, se interpretaba que ya comenzaban los datos;
- si el valor debajo de `"Muestra"` no era numérico, se interpretaba que seguía siendo parte del encabezado.

Ejemplo:

```text
Caso 1: encabezado de una fila

Muestra | MS | PB | FDN
30680   | ...| ...| ...


Caso 2: encabezado de dos filas

Muestra | % | % | %
Nº      | MS| PB| FDN
30680   | ...| ...| ...
```

Esta lógica permitió que el pipeline se adaptara automáticamente a diferentes estructuras sin configurar archivo por archivo.

---

## 7. Normalización de nombres de columnas

Se creó un diccionario llamado `MAPEO_COLUMNAS` para unificar distintas formas históricas de nombrar una misma variable.

Ejemplos:

| Encabezado histórico | Nombre estándar |
|---|---|
| `% NT` | `nt_pct` |
| `% N` | `nt_pct` |
| `% PB` | `pb_pct` |
| `% EE` | `extracto_etereo_pct` |
| `LDAs` | `lignina_pct` |
| `Energía Mcal/kg MS` | `em_mcal_kg_ms` |
| `% Almidón` | `almidon_pct` |

Esto permitió integrar resultados de distintos años bajo una estructura común.

---

## 8. Identificación de muestras reales

Luego de detectar y estandarizar los encabezados, el pipeline debía diferenciar las filas con muestras reales del resto del contenido del informe.

Para eso se utilizó la columna `muestra` como regla principal de validación.

La lógica aplicada fue:

```python
datos["muestra"] = pd.to_numeric(datos["muestra"], errors="coerce")
datos = datos.dropna(subset=["muestra"])
datos["muestra"] = datos["muestra"].astype("Int64")
```

Con esta regla:

- los números de muestra se conservan;
- textos ubicados debajo de la tabla se convierten en `NaN`;
- filas como notas, costos, referencias o pies de informe se eliminan;
- no se eliminan muestras que tengan resultados analíticos faltantes.

Este criterio fue clave para limpiar informes sin borrar información válida.

---

## 9. Conversión de variables analíticas

Las variables analíticas fueron convertidas a formato numérico.

Se aplicaron reglas para:

- reemplazar coma decimal por punto;
- eliminar espacios sobrantes;
- convertir valores a número;
- conservar valores no convertibles como `NaN`.

Un criterio importante fue no reemplazar datos faltantes por cero.

La ausencia de resultado no significa que el resultado sea igual a cero. Por ese motivo, los valores faltantes se conservaron como `NaN`.

---

## 10. Trazabilidad

A cada registro procesado se le incorporó información de trazabilidad.

Entre los campos agregados se incluyeron:

- solicitante;
- localidad;
- fecha de ingreso;
- fecha de egreso;
- archivo de origen.

Esto permitió volver al informe original en caso de detectar alguna inconsistencia durante la validación posterior.

---

## 11. Logs y controles anuales

Cada pipeline anual generó un archivo de log para revisar la calidad del procesamiento.

Los controles incluyeron:

- archivos procesados;
- errores de lectura;
- columnas no mapeadas;
- metadatos faltantes;
- números de muestra duplicados;
- fechas de egreso anteriores a fecha de ingreso;
- fechas fuera del período esperado;
- valores fuera de rango;
- avisos informativos.

El uso de logs permitió validar cada año antes de avanzar al siguiente.

---

## 12. Procesamiento año por año

### 12.1 Año 2018

El procesamiento de 2018 permitió construir la lógica base del pipeline.

Durante esta etapa se resolvieron los principales desafíos estructurales:

- lectura de archivos `.xlsx` y `.xls`;
- detección dinámica de encabezados;
- extracción de metadatos;
- estandarización de columnas;
- filtrado de muestras reales;
- conversión de variables analíticas;
- generación de consolidado anual;
- generación de log de control.

Este año funcionó como base técnica para los años siguientes.

---

### 12.2 Año 2019

Para 2019 se reutilizó la lógica construida en 2018 y se ajustó el pipeline a nuevas estructuras detectadas.

Durante este año aparecieron informes con tablas ubicadas horizontalmente dentro de una misma hoja.

Para resolverlo, se incorporó una lógica especial para procesar bloques horizontales de muestras y luego unificarlos verticalmente.

También se ampliaron reglas de mapeo para nuevas variables históricas.

---

### 12.3 Año 2020

Durante la primera ejecución de 2020 se detectaron tres situaciones particulares:

1. valores múltiples dentro de una misma celda;
2. fechas presentes en el informe pero sin etiqueta identificadora;
3. columnas excluidas que aparecían innecesariamente como avisos.

#### Valores múltiples en una celda

Se detectó una celda con dos determinaciones expresadas en el siguiente formato:

```text
72,15 / 71,57
```

El pipeline esperaba un único valor numérico, por lo que no podía convertir ese texto directamente.

Para resolverlo, se incorporó una función para:

- detectar valores separados por `/`;
- reemplazar coma decimal por punto;
- separar ambas determinaciones;
- convertirlas a número;
- calcular el promedio;
- registrar un aviso en el log.

El resultado consolidado fue el promedio de ambas determinaciones.

Esta solución se incorporó como regla general del pipeline.

#### Corrección documentada de metadatos

También se detectó un caso donde la fecha de ingreso estaba presente en el informe, pero sin la etiqueta esperada.

Como se trataba de una excepción específica y conocida, se incorporó una corrección documentada dentro del pipeline mediante un diccionario de correcciones.

El archivo original no fue modificado.

#### Columnas excluidas

Se ajustó el orden lógico del procesamiento para que una columna definida como excluida no generara avisos innecesarios de columna no mapeada.

Esto permitió mantener logs más limpios y enfocados en situaciones realmente pendientes de revisión.

---

### 12.4 Año 2021

Para 2021 se tomó como base el pipeline corregido de 2020.

La primera ejecución generó avisos relacionados con variantes de una columna denominada `Peso seco`.

Como esa información no formaba parte de las variables analíticas seleccionadas para el proyecto, se decidió excluir sus variantes del consolidado.

La decisión quedó documentada dentro del pipeline mediante `COLUMNAS_EXCLUIDAS`.

Resultado final documentado:

- 38 archivos procesados;
- 236 muestras consolidadas;
- 0 errores de procesamiento;
- 0 fechas de egreso anteriores a la fecha de ingreso;
- 0 valores fuera de los rangos analíticos definidos;
- log final sin avisos pendientes de revisión.

---

### 12.5 Año 2022

Durante el procesamiento de 2022 se detectaron dos situaciones principales.

#### Guion como valor faltante

En varios informes, algunas variables analíticas contenían el valor `-`.

Se interpretó que este guion representaba ausencia de resultado y no un error de formato.

Por eso, el pipeline fue ajustado para convertir estos casos a `NaN` sin generar avisos innecesarios.

#### Tablas horizontales

También se detectó un informe con tres tablas de muestras ubicadas horizontalmente dentro de la misma hoja.

Cada bloque contenía columnas similares, por ejemplo:

```text
Muestra Nº | %NT | %PB
```

El procesamiento estándar interpretaba estas columnas repetidas como duplicadas.

Para resolverlo, se generalizó la función utilizada para tablas horizontales, permitiendo procesar dos o más bloques dentro de una misma hoja.

Los bloques se procesaron por separado y luego se unificaron verticalmente, manteniendo una muestra por fila en el consolidado final.

---

### 12.6 Año 2023

Para 2023 se utilizó como base el pipeline consolidado de 2022.

La ejecución no generó avisos ni errores de procesamiento.

Resultado final documentado:

- 26 archivos procesados;
- 130 muestras consolidadas;
- 0 errores de procesamiento;
- 0 avisos pendientes;
- 0 números de muestra duplicados;
- 0 fechas de egreso anteriores a la fecha de ingreso;
- 0 valores fuera de los rangos analíticos definidos.

---

### 12.7 Año 2024

Para 2024 se reutilizó la lógica validada en 2023.

La primera ejecución generó un aviso correspondiente a la columna `%EE`.

Al revisar la estructura, se identificó que `%EE` correspondía a `Extracto Etéreo`.

Como esta variable ya existía en el modelo bajo el nombre estándar:

```text
extracto_etereo_pct
```

se agregó una nueva equivalencia al diccionario de mapeo:

```python
"ee": "extracto_etereo_pct"
```

Resultado final documentado:

- 32 archivos procesados;
- 147 muestras consolidadas;
- 0 errores de procesamiento;
- log final sin avisos pendientes;
- 0 números de muestra duplicados;
- 0 valores fuera de los rangos analíticos definidos.

---

### 12.8 Año 2025

Para 2025 se tomó como base el pipeline validado de 2024.

La ejecución no generó avisos ni errores de procesamiento.

No fue necesario agregar nuevos mapeos, exclusiones ni tratamientos especiales.

Resultado final documentado:

- 30 archivos procesados;
- 154 muestras consolidadas;
- 0 errores de procesamiento;
- 0 avisos pendientes;
- 0 números de muestra duplicados;
- 0 fechas de egreso anteriores a la fecha de ingreso;
- 0 valores fuera de los rangos analíticos definidos.

---

### 12.9 Año 2026

Para 2026 se utilizó como base el pipeline validado de 2025.

La primera ejecución generó dos tipos de avisos:

1. una columna denominada `% N`;
2. metadatos faltantes en algunos informes.

#### Mapeo de `% N`

Al revisar la estructura, se identificó que `% N` correspondía a nitrógeno total.

Como esa variable ya estaba estandarizada como:

```text
nt_pct
```

se agregó una nueva equivalencia:

```python
"n": "nt_pct"
```

De esta forma, tanto `% NT` como `% N` quedaron integradas en la misma variable estándar.

#### Metadatos faltantes

En algunos informes se detectaron metadatos ausentes, como:

- solicitante;
- localidad;
- fecha de ingreso;
- fecha de egreso.

Al revisar los archivos originales, se comprobó que esa información no estaba disponible.

Por este motivo, se decidió no completar ni inferir estos datos.

Los campos permanecieron como `NaN` y los avisos se conservaron en el log como parte de la trazabilidad.

Este criterio permitió diferenciar:

- un error de procesamiento;
- una limitación real de la información disponible en la fuente original.

---

## 13. Cierre de la etapa histórica 2018–2026

Con la finalización del procesamiento de 2026 quedó completada la normalización anual de informes históricos correspondientes al período 2018–2026.

Durante esta etapa se construyó progresivamente un pipeline capaz de:

- leer archivos `.xlsx` y `.xls`;
- detectar automáticamente encabezados;
- normalizar nombres históricos de variables;
- procesar estructuras de una o múltiples tablas horizontales;
- tratar valores analíticos especiales;
- excluir variables fuera del alcance;
- corregir excepciones documentadas sin modificar archivos originales;
- conservar metadatos y trazabilidad;
- detectar duplicados, fechas inconsistentes y valores fuera de rango;
- registrar avisos y errores en logs.

---

## 14. Script maestro de consolidación histórica

Luego de procesar y validar cada año de forma individual, se desarrolló un script maestro para consolidar los archivos anuales en una única base histórica.

El script maestro transforma:

```text
consolidado_2018.xlsx
consolidado_2019.xlsx
consolidado_2020.xlsx
consolidado_2021.xlsx
consolidado_2022.xlsx
consolidado_2023.xlsx
consolidado_2024.xlsx
consolidado_2025.xlsx
consolidado_2026.xlsx
```

en una base final:

```text
consolidado_historico_2018_2026.xlsx
```

y un archivo de control:

```text
log_consolidado_historico_2018_2026.xlsx
```

---

## 15. Tareas del script maestro

El script maestro realiza las siguientes tareas:

- lee todos los consolidados anuales desde 2018 hasta 2026;
- agrega una columna `anio`;
- unifica todos los archivos mediante Pandas;
- estandariza nombres de columnas;
- ordena las columnas bajo una estructura común;
- convierte fechas a formato fecha;
- formatea fechas visualmente como `DD/MM/YYYY`;
- convierte variables analíticas a formato numérico;
- redondea valores numéricos a dos decimales;
- elimina columnas innecesarias para la base final;
- renombra `observaciones` como `cultivo`;
- estandariza `cultivo` en mayúsculas;
- trata columnas sin nombre para evitar pérdida de información;
- genera controles de calidad históricos;
- exporta una base final y un log.

---

## 16. Decisiones aplicadas en la base histórica final

### 16.1 Eliminación de columnas no necesarias

Para la base final se decidió eliminar columnas utilizadas durante la etapa de trazabilidad interna, como:

- `archivo`;
- `consolidado_origen`.

Estas columnas fueron útiles durante la validación, pero no se consideran necesarias en la base final destinada al análisis.

La trazabilidad general queda conservada en los logs.

---

### 16.2 Renombre de `observaciones` a `cultivo`

La columna `observaciones` fue renombrada como `cultivo`.

Esta decisión se tomó porque, dentro del contexto del proyecto, el contenido de esa columna representaba principalmente la descripción del material analizado, cultivo, muestra o tipo de alimento.

Además, el contenido se estandarizó en mayúsculas.

Ejemplo:

```text
Silo de maíz
silo de maíz
SILO DE MAÍZ
```

queda integrado como:

```text
SILO DE MAÍZ
```

Esto facilita agrupaciones, filtros y análisis posteriores.

---

### 16.3 Estandarización de variables numéricas

Las variables analíticas fueron convertidas a formato numérico y redondeadas a dos decimales.

Algunas de las variables tratadas fueron:

- `ms_pct`;
- `ph`;
- `nt_pct`;
- `pb_pct`;
- `fdn_pct`;
- `fda_pct`;
- `lignina_pct`;
- `divms_pct`;
- `em_mcal_kg_ms`;
- `almidon_pct`;
- `cnes_pct`;
- `extracto_etereo_pct`;
- `ns_nt_pct`;
- `celulosa_pct`.

Las columnas `anio` y `muestra` se mantuvieron como enteros, ya que funcionan como campos identificadores y no como variables analíticas continuas.

---

### 16.4 Corrección final de `celulosa`

Durante la revisión final se detectó que la variable `celulosa` había quedado con un nombre diferente al estándar esperado.

Esto impedía que recibiera el mismo tratamiento de redondeo aplicado al resto de las variables analíticas.

Para resolverlo, se incorporó una regla final:

```text
celulosa → celulosa_pct
```

De esta manera, la variable quedó integrada dentro del conjunto de columnas analíticas y recibió correctamente el formato numérico con dos decimales.

---

### 16.5 Formato de fechas

Las columnas `fecha_ingreso` y `fecha_egreso` se conservaron como fechas reales, no como texto.

El script aplica formato visual en Excel para que se muestren como:

```text
DD/MM/YYYY
```

Esto evita que aparezcan con hora cero y permite que puedan interpretarse correctamente como campos de fecha en análisis posteriores.

---

## 17. Controles históricos generados

El script maestro genera un archivo de log con controles sobre toda la base histórica.

Entre los controles incluidos se encuentran:

- resumen de muestras por año;
- estructura de columnas por consolidado anual;
- valores nulos por columna;
- cobertura de variables analíticas;
- detección de números de muestra duplicados;
- detección de fechas de egreso anteriores a fecha de ingreso;
- detección de fechas fuera del período esperado;
- tratamiento de columnas sin nombre;
- avisos generados durante el proceso;
- errores de lectura o procesamiento.

Estos controles permiten validar la consistencia de la base histórica antes de utilizarla para análisis y visualización.

---

## 18. Resultado final de la etapa

Con esta etapa se obtuvo una base histórica consolidada, estandarizada y lista para análisis:

```text
consolidado_historico_2018_2026.xlsx
```

También se generó un archivo de control:

```text
log_consolidado_historico_2018_2026.xlsx
```

La base final conserva la información analítica de los informes históricos del laboratorio con una estructura homogénea y preparada para una etapa posterior de consulta y visualización.

---

## 19. Aprendizajes técnicos principales

Durante el desarrollo de esta etapa se trabajaron conceptos clave de análisis de datos y construcción de pipelines:

- lectura de archivos Excel con estructuras variables;
- detección dinámica de encabezados;
- normalización de textos;
- diseño de diccionarios de equivalencias;
- validación de datos mediante reglas;
- diferenciación entre dato faltante y valor cero;
- manejo de excepciones documentadas;
- conservación de archivos originales sin modificaciones;
- generación de logs de control;
- consolidación incremental de datos históricos;
- diseño de una base preparada para análisis posterior.

---

## 20. Conclusión

Esta etapa permitió transformar un conjunto de archivos históricos dispersos, con estructuras variables y criterios de carga no estandarizados, en una base histórica consolidada, validada y reutilizable.

El principal resultado no fue solamente un archivo final, sino la construcción de un proceso reproducible para convertir datos operativos desordenados en información estructurada.

La base resultante queda preparada para avanzar hacia la siguiente etapa del proyecto: el diseño de una herramienta de consulta y visualización que facilite el acceso a la información histórica del laboratorio.
