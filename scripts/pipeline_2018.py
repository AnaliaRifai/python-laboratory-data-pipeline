# -*- coding: utf-8 -*-
# ============================================================
# PROYECTO: SISTEMA DE CONSULTA HISTÓRICA - LABORATORIO
# ETAPA: NORMALIZACIÓN DE ARCHIVOS CON PYTHON
# ============================================================
 
# Objetivo:
# Normalizar, validar y consolidar archivos históricos de análisis
# de laboratorio con estructuras heterogéneas en una base anual estandarizada.
# Requisitos del entorno:
# - pandas
# - openpyxl       -> lectura de archivos .xlsx
# - xlrd >= 2.0.1  -> lectura de archivos .xls
# - lxml           -> lectura de .xls que en realidad son HTML
#
# Instalación (desde Anaconda Prompt):
#   conda install xlrd lxml
# Luego reiniciar el kernel de Spyder.
 
 
# ============================================================
# 1. IMPORTACIÓN DE LIBRERÍAS
# ============================================================
 
import pandas as pd
import unicodedata
import re
import glob
import os
 
 
# ============================================================
# 2. CONFIGURACIÓN GENERAL
# ============================================================
 
# Carpeta de entrada: archivos originales de 2018.
# Los archivos originales no se modifican.
CARPETA_2018 = r"C:\Users\54298\Documents\PROGRAMA IPS DATAX\Proyecto Laboratorio\DATOS\ORIGINALES\2018"
 
# Carpeta de salida para guardar puntos de control
# durante el desarrollo del pipeline.
CARPETA_RESPALDOS = r"C:\Users\54298\Documents\PROGRAMA IPS DATAX\Proyecto Laboratorio\DATOS\RESPALDOS"
 
# Nombres de los archivos de salida
ARCHIVO_CONSOLIDADO = "consolidado_2018.xlsx"
ARCHIVO_LOG = "log_consolidado_2018.xlsx"
 
# Muestras que deben excluirse y el motivo.
# Queda documentado aquí para mantener la trazabilidad.
MUESTRAS_EXCLUIDAS = {
    30419: "Registro residual heredado de una plantilla antigua",
}

# Correcciones conocidas de metadatos en archivos históricos
# Se aplican sin modificar los archivos originales.
CORRECCIONES_META = {
    "30781 Criadero el Cencerro.xls": {
        "solicitante": "Criadero el Cencerro"
    },
    "30784 Guillermo Kappes.xls": {
        "solicitante": "Guillermo Kappes"
    }
}

 
# Metadatos que debe tener cada muestra en la base final.
# Si un informe no trae alguno, la columna queda vacía (no falla).
CAMPOS_META = ["solicitante", "localidad", "fecha_ingreso", "fecha_egreso"]

# Variables históricas excluidas del alcance del proyecto
COLUMNAS_EXCLUIDAS = {
    "nida"
}

# Archivos históricos duplicados u obsoletos
# Se excluyen porque existe una versión más completa o
# porque contienen registros acumulados ya presentes
# en informes individuales.
ARCHIVOS_EXCLUIDOS = {
    "30762 Ayelén.xls":
        "Versión incompleta; existe un informe posterior más completo de la misma muestra",

    "30798 Gerardo Carreira - Bonifacio.xls":
        "Informe acumulativo que repite muestras presentes en informes individuales"
}
 
# Rangos válidos esperados para cada variable analítica.
# Se usan en el control de calidad para detectar valores anómalos.
RANGOS_VALIDOS = {
    "ms_pct": (0, 100),
    "ph": (0, 14),
    "nt_pct": (0, 100),
    "pb_pct": (0, 100),
    "fdn_pct": (0, 100),
    "fda_pct": (0, 100),
    "lignina_pct": (0, 100),
    "divms_pct": (0, 100),
    "em_mcal_kg_ms": (0, 5),
    "almidon_pct": (0, 100),
    "cnes_pct": (0, 100),
}
 
 
# ============================================================
# 3. ESTANDARIZACIÓN DE NOMBRES DE COLUMNAS
# ============================================================
# Diccionario de equivalencias:
# clave = nombre limpio detectado en el archivo
# valor = nombre estándar que tendrá en la base consolidada
#
# Nota: la columna de muestra no depende de este diccionario,
# se nombra directamente por su posición (ver procesar_reporte).
 
MAPEO_COLUMNAS = {
    "muestrano": "muestra",
    "muestra": "muestra",
    "ms": "ms_pct",
    "ph": "ph",
    "nt": "nt_pct",
    "pb": "pb_pct",
    "fdn": "fdn_pct",
    "fda": "fda_pct",
    "lignina": "lignina_pct",
    "ldas": "lignina_pct",
    "divms": "divms_pct",
    "divmsmcalkgms": "divms_pct",
    "emmcalkgms": "em_mcal_kg_ms",
    "energiamcalkgms": "em_mcal_kg_ms",
    "energia": "em_mcal_kg_ms",
    "actividadureasica": "actividad_ureasica",
    "almidon": "almidon_pct",
    "cnes": "cnes_pct",
    "observaciones": "observaciones",
}
 
 
# ============================================================
# 4. REGISTRO DE AVISOS
# ============================================================
# Todos los avisos se guardan en una lista para exportarlos
# al final como log. Con muchos archivos, la consola no alcanza.
 
registro_avisos = []
 
 
def registrar_aviso(archivo, tipo, detalle):
    """Guarda un aviso en el registro y lo muestra en consola."""
    registro_avisos.append({
        "archivo": archivo,
        "tipo": tipo,
        "detalle": str(detalle),
    })
    print(f"  Aviso [{tipo}] {archivo}: {detalle}")
 
 
# ============================================================
# 5. FUNCIONES AUXILIARES
# ============================================================
 
def limpiar_texto(valor):
    """
    Normaliza un texto para poder compararlo aunque tenga
    diferencias de formato.
 
    La función:
    - convierte el valor a texto;
    - elimina tildes y caracteres especiales;
    - convierte todo a minúsculas;
    - elimina espacios, símbolos y signos.
 
    Ejemplos:
    '% Almidón' -> 'almidon'
    'Muestra Nº' -> 'muestrano'
    'ENERGIA Mcal/kg MS' -> 'energiamcalkgms'
    """
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", texto.lower())
 
 
def es_numero(valor):
    """Devuelve True si el valor puede interpretarse como número."""
    return pd.notna(pd.to_numeric(valor, errors="coerce"))
 
 
def leer_archivo(ruta):
    """
    Lee un archivo Excel completo, sin asumir dónde está el encabezado.
 
    - Los .xlsx se leen con openpyxl y los .xls con xlrd
      (pandas elige el motor según la extensión).
    - Algunos sistemas de laboratorio exportan archivos HTML con
      extensión .xls. Si la lectura como Excel falla, se intenta
      leerlo como tabla HTML y se deja registrado en el log.
    """
    try:
        return pd.read_excel(ruta, header=None)
 
    except ImportError:
        # Falta una librería (por ejemplo xlrd): no tiene sentido
        # intentar otra lectura, hay que instalarla.
        raise
 
    except Exception as error_excel:
        # Solo se intenta la lectura alternativa en archivos .xls
        if not ruta.lower().endswith(".xls"):
            raise
 
        try:
            # thousands=None evita que "12,5" se interprete como 125
            tabla = pd.read_html(ruta, header=None, thousands=None)[0]
        except Exception:
            raise ValueError(
                f"No se pudo leer como Excel ni como HTML ({error_excel})"
            )
 
        registrar_aviso(
            os.path.basename(ruta),
            "formato",
            "archivo .xls leído como HTML"
        )
        return tabla
 
 
def buscar_encabezado(df):
    """
    Busca automáticamente dónde comienza la tabla de muestras.
 
    Recorre el DataFrame fila por fila y celda por celda hasta
    encontrar una celda que EMPIECE con 'muestra' y que no sea
    una etiqueta de metadato (no contiene ':').
 
    Esto evita falsos positivos como 'Tipo de muestra: Ensilaje'
    o 'Muestras recibidas: 12'.
 
    Devuelve:
    - i: posición de la fila donde se encontró el encabezado.
    - j: posición de la columna donde se encontró 'muestra'.
    """
 
    # Recorre todas las filas del DataFrame
    for i in range(len(df)):
 
        # enumerate() entrega la posición de la columna (j)
        # y el contenido de la celda (valor)
        for j, valor in enumerate(df.iloc[i].values):
 
            if (
                isinstance(valor, str)
                and limpiar_texto(valor).startswith("muestra")
                and ":" not in valor
            ):
                return i, j
 
    # Si se recorrió todo el archivo sin encontrar el encabezado,
    # se detiene el procesamiento de ese archivo con un error.
    raise ValueError("No se encontró la fila de encabezado")
 
 
def extraer_metadatos(df, fila_encabezado):
    """
    Extrae la información general del informe ubicada antes
    de la tabla de muestras.
 
    Busca: solicitante, localidad, fecha de ingreso y fecha de egreso.
 
    Soporta dos formatos:
    - Etiqueta y valor en la misma celda: 'Solicitante: Juan Pérez'
    - Etiqueta y valor en celdas contiguas: 'Solicitante:' | 'Juan Pérez'
 
    IMPORTANTE: debe recibir el DataFrame ORIGINAL (sin recortar),
    porque los metadatos pueden estar a la izquierda de la tabla.
 
    Devuelve un diccionario con los metadatos encontrados.
    """
 
    # Claves normalizadas con limpiar_texto (tolera tildes y mayúsculas)
    claves = {
        "solicitante": "solicitante",
        "localidad": "localidad",
        "fechadeingreso": "fecha_ingreso",
        "fechadeegreso": "fecha_egreso",
    }
 
    meta = {}
 
    # Recorre solo las filas anteriores al encabezado
    for _, fila_df in df.iloc[:fila_encabezado].iterrows():
        valores = fila_df.tolist()
 
        for j, valor in enumerate(valores):
 
            # Solo interesan textos con formato 'Etiqueta: ...'
            if not isinstance(valor, str) or ":" not in valor:
                continue
 
            etiqueta, _, contenido = valor.partition(":")
            clave = limpiar_texto(etiqueta)
 
            if clave not in claves:
                continue
 
            contenido = contenido.strip()
 
            # Si el valor no está en la misma celda,
            # se toma la siguiente celda no vacía de la fila
            if not contenido:
                siguientes = [
                    v for v in valores[j + 1:]
                    if pd.notna(v) and str(v).strip() != ""
                ]
                contenido = siguientes[0] if siguientes else None
 
            meta[claves[clave]] = contenido
 
    return meta
 
 
# ============================================================
# 6. PROCESAMIENTO DE UN INFORME
# ============================================================
 
def procesar_reporte(ruta):
    """
    Procesa un archivo individual del laboratorio y devuelve
    una tabla estandarizada con sus muestras.
    """
    nombre_archivo = os.path.basename(ruta)
 
    # 1) Leer el archivo y ubicar el encabezado
    # Se conserva una copia intacta para extraer los metadatos
    df_original = leer_archivo(ruta)
    fila, col = buscar_encabezado(df_original)
 
    # Trabajar la tabla de muestras sobre una copia recortada,
    # descartando las columnas ubicadas a la izquierda de "Muestra"
    df = df_original.iloc[:, col:].copy()
 
    # Reiniciar temporalmente los nombres de las columnas
    df.columns = range(df.shape[1])
 
    # 2) Detectar si el encabezado ocupa una o dos filas
    # Si la fila siguiente no empieza con un número, es parte del encabezado
    dos_filas = (fila + 1 < len(df)) and not es_numero(df.iloc[fila + 1, 0])
    filas_enc = [fila, fila + 1] if dos_filas else [fila]
 
    # 3) Construir y estandarizar los nombres de las columnas
    nombres = []
 
    for c in df.columns:
        # Une el texto de las filas de encabezado (una o dos)
        partes = [
            str(df.iloc[f, c]).strip()
            for f in filas_enc
            if pd.notna(df.iloc[f, c]) and str(df.iloc[f, c]).strip() != ""
        ]
 
        nombre_crudo = " ".join(partes)
        clave = limpiar_texto(nombre_crudo)
 
        if clave == "":
            nombres.append(f"sin_nombre_{c}")
        else:
            # La columna 0 se nombra aparte, no requiere aviso
            if clave not in MAPEO_COLUMNAS and c != 0:
                registrar_aviso(nombre_archivo, "columna no mapeada", nombre_crudo)
 
            nombres.append(MAPEO_COLUMNAS.get(clave, clave))
 
    # La columna 0 es la que contenía "muestra" (se recortó desde ahí),
    # así que se nombra directamente sin depender del mapeo.
    # Esto cubre variantes como 'Muestra N°', 'Muestra Nº', 'N° Muestra', etc.
    nombres[0] = "muestra"
 
    # 4) Evitar nombres de columnas duplicados
    vistos = {}
 
    for i, n in enumerate(nombres):
        vistos[n] = vistos.get(n, 0) + 1
 
        if vistos[n] > 1:
            nombres[i] = f"{n}_{vistos[n]}"
 
            # Las repeticiones de DIVMS se esperan y se promedian más abajo;
            # cualquier otro duplicado se informa para revisarlo
            if not n.startswith("divms_pct"):
                registrar_aviso(nombre_archivo, "columna duplicada", n)
 
    # 5) Extraer las filas correspondientes a las muestras
    # .copy() crea un DataFrame independiente llamado datos
    datos = df.iloc[filas_enc[-1] + 1:].copy()
 
    # Asignar los nombres de columnas ya estandarizados
    datos.columns = nombres
    
    # Eliminar variables históricas fuera del alcance del proyecto
    columnas_a_eliminar = [
        c for c in datos.columns
        if c in COLUMNAS_EXCLUIDAS
    ]

    datos = datos.drop(columns=columnas_a_eliminar)
 
    # Convertir el número de muestra a formato numérico
    datos["muestra"] = pd.to_numeric(datos["muestra"], errors="coerce")
 
    # Eliminar filas que no correspondan a muestras reales
    # (filas vacías, totales, firmas, notas al pie, etc.)
    datos = datos.dropna(subset=["muestra"]).copy()
 
    # Guardar el número de muestra como entero
    datos["muestra"] = datos["muestra"].astype("Int64")
 
    # Excluir muestras definidas en MUESTRAS_EXCLUIDAS y dejarlo registrado
    excluidas = datos["muestra"].isin(list(MUESTRAS_EXCLUIDAS))
 
    for m in datos.loc[excluidas, "muestra"]:
        registrar_aviso(
            nombre_archivo,
            "muestra excluida",
            f"{m}: {MUESTRAS_EXCLUIDAS[int(m)]}"
        )
 
    datos = datos[~excluidas].copy()
 
    # 6) Convertir las variables analíticas a formato numérico
    for c in datos.columns:
        if c in ("muestra", "observaciones"):
            continue
 
        original = datos[c]
 
        # Reemplaza coma decimal por punto y convierte a número
        convertido = pd.to_numeric(
            original.astype(str).str.replace(",", ".").str.strip(),
            errors="coerce"
        )
 
        # Valores que tenían contenido pero no se pudieron convertir
        # (por ejemplo '<0,1', 'ND', 'Trazas'). Se informan para no
        # perder información en silencio.
        perdidos = original.notna() & convertido.isna()
 
        if perdidos.any():
            ejemplos = original[perdidos].astype(str).unique()[:5].tolist()
            registrar_aviso(
                nombre_archivo,
                "valor no numérico",
                f"{perdidos.sum()} valores en '{c}': {ejemplos}"
            )
 
        datos[c] = convertido
 
    # 7) Promediar repeticiones de DIVMS cuando existan
    # Cubre cualquier cantidad: divms_pct, divms_pct_2, divms_pct_3...
    cols_divms = [c for c in datos.columns if c.startswith("divms_pct")]
 
    if len(cols_divms) > 1:
        datos["divms_pct"] = datos[cols_divms].mean(axis=1)
        datos = datos.drop(columns=[c for c in cols_divms if c != "divms_pct"])
 
    # 8) Eliminar columnas completamente vacías
    datos = datos.dropna(axis=1, how="all")
 
    # 9) Agregar metadatos del informe a cada muestra
    # Se extraen del archivo ORIGINAL, no del recortado
    meta = extraer_metadatos(df_original, fila)
    
    # Aplicar correcciones conocidas de metadatos
    if nombre_archivo in CORRECCIONES_META:
        meta.update(CORRECCIONES_META[nombre_archivo])
 
    # Garantiza que todas las columnas existan, aunque el informe
    # no traiga el dato (así el archivo no se descarta por un KeyError)
    for campo in CAMPOS_META:
        datos[campo] = meta.get(campo)
 
        if campo not in meta:
            registrar_aviso(nombre_archivo, "metadato faltante", campo)
 
    # Convertir las fechas a formato fecha (día/mes/año)
    for campo in ("fecha_ingreso", "fecha_egreso"):
        datos[campo] = pd.to_datetime(datos[campo], dayfirst=True, errors="coerce")
 
    # Agregar el nombre del archivo de origen
    datos["archivo"] = nombre_archivo
 
    # 10) Devolver el DataFrame procesado
    return datos
 
 
# ============================================================
# 7. BÚSQUEDA DE ARCHIVOS DE 2018
# ============================================================
 
archivos_2018 = (
    glob.glob(os.path.join(CARPETA_2018, "*.xlsx")) +
    glob.glob(os.path.join(CARPETA_2018, "*.xls"))
)
 
# Excluir archivos temporales generados por Excel (~$)
archivos_2018 = [
    ruta for ruta in archivos_2018
    if not os.path.basename(ruta).startswith("~$")
]
 
print("Cantidad de archivos encontrados:", len(archivos_2018))
print("  .xlsx:", sum(r.lower().endswith(".xlsx") for r in archivos_2018))
print("  .xls :", sum(r.lower().endswith(".xls") for r in archivos_2018))
 
# 8. CONSOLIDACIÓN DE TODOS LOS ARCHIVOS (.XLSX Y .XLS)
# ============================================================
 
# Lista donde se guardará el DataFrame de cada informe
tablas = []
 
# Lista donde se guardarán los archivos que no se pudieron procesar
registro_errores = []
 
# Procesar cada archivo
for ruta in sorted(archivos_2018):
    nombre_archivo = os.path.basename(ruta)
    
    if nombre_archivo in ARCHIVOS_EXCLUIDOS:
        registrar_aviso(
            nombre_archivo,
            "archivo excluido",
            ARCHIVOS_EXCLUIDOS[nombre_archivo]
        )
        continue
 
    try:
        tabla = procesar_reporte(ruta)
        tablas.append(tabla)
        print("OK:", nombre_archivo, "| Muestras:", len(tabla))
 
    except Exception as e:
        registro_errores.append({"archivo": nombre_archivo, "error": str(e)})
        print("ERROR:", nombre_archivo, "->", e)
 
# Si no se procesó ningún archivo, se detiene el script
if not tablas:
    raise SystemExit("No se pudo procesar ningún archivo. Revisar los errores.")
 
# Unir todos los DataFrames en una única tabla
consolidado = pd.concat(tablas, ignore_index=True)
 
# Ordenar columnas: identificación y metadatos primero, luego analíticas
columnas_inicio = ["archivo", "muestra"] + CAMPOS_META
consolidado = consolidado[
    columnas_inicio
    + [c for c in consolidado.columns if c not in columnas_inicio]
]
 
print("\n--- RESUMEN DEL CONSOLIDADO ---")
print("Archivos encontrados:", len(archivos_2018))
print("Archivos procesados:", len(tablas))
print("Archivos con error:", len(registro_errores))
print("Muestras totales:", len(consolidado))
print("Columnas:", consolidado.columns.tolist())
 
 
# ============================================================
# 9. CONTROLES DE CALIDAD
# ============================================================
 
# ---------- 9.1 Duplicados ----------
# Un número de muestra repetido puede indicar un informe
# guardado dos veces o una numeración reutilizada.
duplicados = consolidado[
    consolidado.duplicated(subset=["muestra"], keep=False)
]
 
print("\n--- CONTROL DE DUPLICADOS ---")
print("Filas con número de muestra duplicado:", len(duplicados))
 
if not duplicados.empty:
    print(
        duplicados[["muestra", "solicitante", "fecha_ingreso", "archivo"]]
        .sort_values("muestra")
    )
 
# ---------- 9.2 Valores faltantes ----------
print("\n--- CONTROL DE VALORES FALTANTES ---")
print(consolidado.isna().sum())
 
# ---------- 9.3 Tipos de datos ----------
print("\n--- CONTROL DE TIPOS DE DATOS ---")
print(consolidado.dtypes)
 
# ---------- 9.4 Fechas ----------
print("\n--- CONTROL DE FECHAS ---")
 
# Fechas que no se pudieron interpretar
print("Fechas de ingreso vacías:", consolidado["fecha_ingreso"].isna().sum())
print("Fechas de egreso vacías:", consolidado["fecha_egreso"].isna().sum())
 
# Rango de fechas: todo debería estar dentro de 2018
print("Rango fecha_ingreso:",
      consolidado["fecha_ingreso"].min(), "->",
      consolidado["fecha_ingreso"].max())
 
# Egreso anterior al ingreso: indica un error de digitación
fechas_invertidas = consolidado[
    consolidado["fecha_egreso"] < consolidado["fecha_ingreso"]
]
print("Muestras con egreso anterior al ingreso:", len(fechas_invertidas))
 
# ---------- 9.5 Rangos de variables analíticas ----------
print("\n--- CONTROL DE RANGOS ANALÍTICOS ---")
 
for columna, (minimo, maximo) in RANGOS_VALIDOS.items():
    if columna not in consolidado.columns:
        continue
 
    serie = consolidado[columna]
 
    # Valores con dato que quedan fuera del rango esperado
    fuera = consolidado[serie.notna() & ~serie.between(minimo, maximo)]
 
    print(
        f"{columna:<15} | Mín: {serie.min()} | Máx: {serie.max()} "
        f"| Fuera de rango ({minimo}-{maximo}): {len(fuera)}"
    )
 
    if not fuera.empty:
        print(fuera[["archivo", "muestra", columna]].to_string(index=False))
 
# ---------- 9.6 Escala por archivo ----------
# Detecta archivos que informan en otra escala
# (por ejemplo MS = 0.35 en lugar de 35)
print("\n--- CONTROL DE ESCALA POR ARCHIVO (medianas) ---")
 
columnas_escala = [
    c for c in ["ms_pct", "pb_pct", "fdn_pct", "fda_pct"]
    if c in consolidado.columns
]
 
print(consolidado.groupby("archivo")[columnas_escala].median())
 
# ---------- 9.7 Trazabilidad ----------
muestras_por_archivo = (
    consolidado
    .groupby("archivo")
    .size()
    .sort_index()
)
 
print("\n--- CONTROL FINAL DE TRAZABILIDAD ---")
print("Archivos procesados:", consolidado["archivo"].nunique())
print("Muestras totales:", len(consolidado))
print("\nMuestras por archivo:")
print(muestras_por_archivo)
 
 
# ============================================================
# 10. EXPORTACIÓN
# ============================================================
 
# Crear la carpeta de salida si no existe
os.makedirs(CARPETA_RESPALDOS, exist_ok=True)
 
# Consolidado anual limpio y estandarizado (una sola hoja)
ruta_consolidado = os.path.join(CARPETA_RESPALDOS, ARCHIVO_CONSOLIDADO)
consolidado.to_excel(ruta_consolidado, index=False)
 
# Log: avisos, errores y muestras por archivo en un archivo aparte
ruta_log = os.path.join(CARPETA_RESPALDOS, ARCHIVO_LOG)
 
with pd.ExcelWriter(ruta_log) as writer:
    pd.DataFrame(registro_avisos, columns=["archivo", "tipo", "detalle"]) \
        .to_excel(writer, sheet_name="avisos", index=False)
 
    pd.DataFrame(registro_errores, columns=["archivo", "error"]) \
        .to_excel(writer, sheet_name="errores", index=False)
 
    muestras_por_archivo.rename("muestras").to_frame() \
        .to_excel(writer, sheet_name="muestras_por_archivo")
 
print("\n--- EXPORTACIÓN FINALIZADA ---")
print("Consolidado:", ruta_consolidado)
print("Log:", ruta_log)
