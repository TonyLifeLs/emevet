"""
Genera PrincipioActivoAnalisisMercado2026.xlsx consolidando:

- Todos los archivos Excel dentro de la carpeta "Mercado/" (fuente de presentaciones)
- Todos los archivos Reporte_Por_Principio_Activo_2024_*.xlsx en el directorio raíz
  (fuente de ComponenteMolecular y datos de ventas por categoría)

Lógica de enriquecimiento (como químico):
1. Para cada Presentacion se extrae el PrincipioActivo y se busca su ComponenteMolecular
   en los archivos de Reporte mediante normalización química (acentos, prefijos
   estereoquímicos, sufijos de sales, equivalencias inglés/español).
2. Propagación: si un PrincipioActivo ya tiene ComponenteMolecular en alguna fila,
   ese valor se propaga a las demás presentaciones del mismo PA que no tengan uno.
3. Simulación: si tras la búsqueda y propagación un PA sigue sin ComponenteMolecular,
   se genera un valor coherente basado en el nombre del PA (diccionario de drogas
   conocidas o "COMPUESTO GENERICO_" + nombre).
4. El campo final se llama ComponenteMolecularFinal y NUNCA queda vacío.
"""

import glob
import os
import re
import unicodedata

import pandas as pd


# ---------------------------------------------------------------------------
# Utilidades de normalización química
# ---------------------------------------------------------------------------

def quitar_acentos(texto: str) -> str:
    """Elimina tildes/diacríticos manteniendo la letra base."""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# Equivalencias ortográficas: inglés / variantes → español canónico
EQUIVALENCIAS: dict[str, str] = {
    "METHIONINA": "METIONINA",
    "METHIONINE": "METIONINA",
    "LYSINE": "LISINA",
    "THREONINE": "TREONINA",
    "TRYPTOPHAN": "TRIPTOFANO",
    "PHENYLALANINE": "FENILALANINA",
    "VALINE": "VALINA",
    "LEUCINE": "LEUCINA",
    "ISOLEUCINE": "ISOLEUCINA",
    "ARGININE": "ARGININA",
    "GLYCINE": "GLICINA",
    "ALANINE": "ALANINA",
    "SERINE": "SERINA",
    "CYSTINE": "CISTINA",
    "HISTIDINE": "HISTIDINA",
    "PROLINE": "PROLINA",
    "TYROSINE": "TIROSINA",
    "GLUTAMINE": "GLUTAMINA",
    "ASPARAGINE": "ASPARAGINA",
    "CHOLINE": "COLINA",
    "BETAINE": "BETAINA",
    "IVERMECTIN": "IVERMECTINA",
    "ERYTHROMYCIN": "ERITROMICINA",
    "DOXYCYCLINE": "DOXICICLINA",
    "OXYTETRACYCLINE": "OXITETRACICLINA",
    "TETRACYCLINE": "TETRACICLINA",
    "AMPICILLIN": "AMPICILINA",
    "AMOXICILLIN": "AMOXICILINA",
    "GENTAMICIN": "GENTAMICINA",
    "NEOMYCIN": "NEOMICINA",
    "PENICILLIN": "PENICILINA",
    "ENROFLOXACIN": "ENROFLOXACINA",
    "CIPROFLOXACIN": "CIPROFLOXACINA",
    "NORFLOXACIN": "NORFLOXACINA",
    "TILMICOSIN": "TILMICOSINA",
    "TYLOSIN": "TILOSINA",
    "LINCOMYCIN": "LINCOMICINA",
    "SPECTINOMYCIN": "ESPECTINOMICINA",
    "COLISTIN": "COLISTINA",
    "DEXAMETHASONE": "DEXAMETASONA",
    "PREDNISOLONE": "PREDNISOLONA",
    "KETAMINE": "KETAMINA",
    "XYLAZINE": "XILAZINA",
    "BUTORPHANOL": "BUTORFANOL",
    "OXYTOCIN": "OXITOCINA",
    "PROGESTERONE": "PROGESTERONA",
    "TESTOSTERONE": "TESTOSTERONA",
    "SELENIUM": "SELENIO",
    "VITAMIN": "VITAMINA",
    "VITAMINS": "VITAMINAS",
    "ALBENDAZOLE": "ALBENDAZOL",
    "FENBENDAZOLE": "FENBENDAZOL",
    "PRAZIQUANTEL": "PRAZICUANTEL",
    "DORAMECTIN": "DORAMECTINA",
    "ABAMECTIN": "ABAMECTINA",
    "MOXIDECTIN": "MOXIDECTINA",
    "EPRINOMECTIN": "EPRINOMECTINA",
    "FLUBENDAZOLE": "FLUBENDAZOL",
    "LEVAMISOLE": "LEVAMISOL",
    "NICLOSAMIDE": "NICLOSAMIDA",
    "PYRANTEL": "PIRANTEL",
    "ZINC": "ZINC",
    "COPPER": "COBRE",
    "MANGANESE": "MANGANESO",
    "IRON": "HIERRO",
    "CALCIUM": "CALCIO",
    "PHOSPHORUS": "FOSFORO",
    "SODIUM": "SODIO",
    "POTASSIUM": "POTASIO",
    "MAGNESIUM": "MAGNESIO",
    "IODINE": "YODO",
}

# Palabras calificadoras que no forman parte del nombre molecular
_CALIFICADORAS = (
    r"AVICOLA|AVIAR|BOVINO|BOVINA|CANINO|CANINA|FELINO|FELINA|EQUINO|EQUINA|"
    r"PORCINO|PORCINA|OVINO|OVINA|CAPRINO|CAPRINA|CUNICULAR|"
    r"INYECTABLE|ORAL|TOPICO|TOPICA|INTRAVENOSO|SUBCUTANEO|"
    r"FORTE|PLUS|ULTRA|SUPER|MAX|PRO|PREMIUM|EXTRA|SPECIAL|"
    r"PECUARIO|VETERINARIO|VETERINARIA"
)
PALABRAS_NO_MOLECULARES = re.compile(rf"\b({_CALIFICADORAS})\b", re.IGNORECASE)

# Prefijos de marca cortos al inicio del texto
PREFIJOS_MARCA = re.compile(r"^(LH|VM|BM|AG|QSI|MC)\s*", re.IGNORECASE)

# Prefijos estereoquímicos que no cambian la identidad de la molécula
PREFIJOS_QUIRALES = re.compile(
    r"^(DL|L|D|R|S|MESO|CIS|TRANS|ALFA|BETA|GAMMA|DELTA)-\s*", re.IGNORECASE
)

# Sufijos de sales / formas farmacéuticas
SUFIJOS_SALES = re.compile(
    r"\s+(HCL|CLORHIDRATO|HIDROCLORURO|SULFATO|FOSFATO|TARTRATO|CITRATO|LACTATO|"
    r"GLUCONATO|ACETATO|NITRATO|BROMURO|YODURO|MALEATO|FUMARATO|SUCCINATO|"
    r"SODICO|POTASICO|MAGNESICO|AMONICO|SODICO|POTASICO|"
    r"TRIHIDRATADO|TRIHIDRATADA|MONOHIDRATADO|HIDRATADO|ANHIDRO|BASE|"
    r"MICRONIZADO|MICRONIZADA|GRANULADO|GRANULADA|POLVO)$",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Diccionario de simulación química
# Patrón en el nombre del PA → ComponenteMolecularFinal simulado
# ---------------------------------------------------------------------------
SIMULACION_DROGAS: list[tuple[str, str]] = [
    # Antiparasitarios
    (r"IVERMECT",       "IVERMECTINA"),
    (r"DORAMECT",       "DORAMECTINA"),
    (r"ABAMECT",        "ABAMECTINA"),
    (r"MOXIDECT",       "MOXIDECTINA"),
    (r"EPRINOMECT",     "EPRINOMECTINA"),
    (r"ALBENDAZ",       "ALBENDAZOL"),
    (r"FENBENDAZ",      "FENBENDAZOL"),
    (r"FLUBENDAZ",      "FLUBENDAZOL"),
    (r"OXIBENDAZ",      "OXIBENDAZOL"),
    (r"MEBENDAZ",       "MEBENDAZOL"),
    (r"PRAZIQUANT",     "PRAZICUANTEL"),
    (r"LEVAMIS",        "LEVAMISOL"),
    (r"NICLOSAMID",     "NICLOSAMIDA"),
    (r"PIRANTEL",       "PIRANTEL"),
    (r"CLOSANTEL",      "CLOSANTEL"),
    (r"TRICLABENDAZ",   "TRICLABENDAZOL"),
    (r"MONENSINA",      "MONENSINA SODICA"),
    (r"LASALOCID",      "LASALOCID SODICO"),
    (r"SALINOMICIN",    "SALINOMICINA"),
    (r"NARASIN",        "NARASINA"),
    (r"ROBENIDINA",     "ROBENIDIN"),
    (r"DICLAZURIL",     "DICLAZURIL"),
    (r"TOLTRAZURIL",    "TOLTRAZURIL"),
    (r"AMPROLIUM",      "AMPROLIUM"),
    # Antibióticos
    (r"ENROFLOXAC",     "ENROFLOXACINA"),
    (r"CIPROFLOXAC",    "CIPROFLOXACINA"),
    (r"NORFLOXAC",      "NORFLOXACINA"),
    (r"DANOFLOXAC",     "DANOFLOXACINA"),
    (r"MARBOFLOXAC",    "MARBOFLOXACINA"),
    (r"FLORFENICOB?",    "FLORFENICOL"),
    (r"OXITETRACICL",   "OXITETRACICLINA"),
    (r"DOXICICL",       "DOXICICLINA"),
    (r"TETRACICLINA",   "TETRACICLINA"),
    (r"CLORTETRACICLINA","CLORTETRACICLINA"),
    (r"TILVALOSINA",    "TILVALOSINA"),
    (r"TILMICOSINA",    "TILMICOSINA"),
    (r"TILOSINA",       "TILOSINA TARTRATO"),
    (r"ERITROMICIN",    "ERITROMICINA"),
    (r"LINCOMICIN",     "LINCOMICINA"),
    (r"ESPECTINOMICIN", "ESPECTINOMICINA"),
    (r"NEOMICIN",       "NEOMICINA SULFATO"),
    (r"GENTAMICIN",     "GENTAMICINA SULFATO"),
    (r"AMIKACIN",       "AMIKACINA"),
    (r"ESTREPTOMICIN",  "ESTREPTOMICINA"),
    (r"DIHIDROESTREPT", "DIHIDROESTREPTOMICINA"),
    (r"PENICILINA",     "PENICILINA G PROCAINA"),
    (r"PENICILLI",      "PENICILINA G PROCAINA"),
    (r"AMPICILI",       "AMPICILINA TRIHIDRATADA"),
    (r"AMOXICILI",      "AMOXICILINA TRIHIDRATADA"),
    (r"CLOXA",          "CLOXACILINA SODICA"),
    (r"CEFTIOFUR",      "CEFTIOFUR CLORHIDRATO"),
    (r"CEFADROXIL",     "CEFADROXIL"),
    (r"COLISTIN",       "COLISTINA SULFATO"),
    (r"AVILAMICIN",     "AVILAMICINA"),
    (r"SULFADIAZIN",    "SULFADIAZINA"),
    (r"SULFAMETAZIN",   "SULFAMETAZINA"),
    (r"SULFADIMIDINA",  "SULFADIMIDINA"),
    (r"TRIMETO",        "TRIMETOPRIMA"),
    # AINEs / Analgésicos
    (r"MELOXICAM",      "MELOXICAM"),
    (r"CARPROFENO",     "CARPROFENO"),
    (r"DICLOFENAC",     "DICLOFENACO SODICO"),
    (r"FLUNIXIN",       "FLUNIXIN MEGLUMINE"),
    (r"FENILBUTAZON",   "FENILBUTAZONA"),
    (r"KETOPROFENO",    "KETOPROFENO"),
    (r"IBUPROFENO",     "IBUPROFENO"),
    (r"PARACETAMOL",    "ACETAMINOFEN"),
    (r"ACETAMINOFEN",   "ACETAMINOFEN"),
    # Corticoides / Hormonas
    (r"DEXAMETASON",    "DEXAMETASONA"),
    (r"PREDNISOLON",    "PREDNISOLONA"),
    (r"CORTISOL",       "HIDROCORTISONA"),
    (r"BETAMETASON",    "BETAMETASONA"),
    (r"PROGESTERON",    "PROGESTERONA"),
    (r"TESTOSTERON",    "TESTOSTERONA PROPIONATO"),
    (r"ESTRADIOL",      "ESTRADIOL BENZOATO"),
    (r"OXITOCINA",      "OXITOCINA"),
    (r"OXYTOCIN",       "OXITOCINA"),
    (r"GONADOTROP",     "GONADOTROPINA CORIONICA"),
    (r"FSH",            "HORMONA FOLICULOESTIMULANTE"),
    (r"GNRH|GONADOR",   "GONADORELINA"),
    (r"INSULINA",       "INSULINA"),
    (r"SOMATOTROPIN",   "SOMATOTROPINA BOVINA"),
    # Anestésicos / Sedantes
    (r"KETAMINA",       "KETAMINA CLORHIDRATO"),
    (r"XILAZINA",       "XILAZINA CLORHIDRATO"),
    (r"BUTORFANOL",     "BUTORFANOL TARTRATO"),
    (r"ROMIFIDINA",     "ROMIFIDINA"),
    (r"DETOMIDINA",     "DETOMIDINA"),
    (r"MEDETOMIDINA",   "MEDETOMIDINA"),
    (r"PROPOFOL",       "PROPOFOL"),
    (r"TILETAMINA",     "TILETAMINA"),
    (r"ZOLAZEPAM",      "ZOLAZEPAM"),
    # Vitaminas y aminoácidos
    (r"METIONINA|METHIONIN", "METIONINA"),
    (r"LISINA",         "LISINA"),
    (r"TREONINA",       "TREONINA"),
    (r"TRIPTOFANO",     "TRIPTOFANO"),
    (r"ARGININA",       "ARGININA"),
    (r"VALINA",         "VALINA"),
    (r"LEUCINA",        "LEUCINA"),
    (r"ISOLEUCINA",     "ISOLEUCINA"),
    (r"GLICINA",        "GLICINA"),
    (r"COLINA",         "CLORURO DE COLINA"),
    (r"VITAMINA A",     "VITAMINA A (RETINOL)"),
    (r"VITAMINA D",     "VITAMINA D3 (COLECALCIFEROL)"),
    (r"VITAMINA E",     "VITAMINA E (ALFA-TOCOFEROL)"),
    (r"VITAMINA K",     "VITAMINA K3 (MENADIONA)"),
    (r"VITAMINA C",     "ACIDO ASCORBICO"),
    (r"VITAMINA B1",    "TIAMINA"),
    (r"VITAMINA B2",    "RIBOFLAVINA"),
    (r"VITAMINA B6",    "PIRIDOXINA"),
    (r"VITAMINA B12",   "CIANOCOBALAMINA"),
    (r"BIOTINA",        "BIOTINA"),
    (r"ACIDO FOLICO",   "ACIDO FOLICO"),
    (r"NIACINA",        "NIACINA"),
    # Minerales
    (r"CALCIO",         "CALCIO"),
    (r"FOSFORO",        "FOSFORO"),
    (r"MAGNESIO",       "MAGNESIO"),
    (r"ZINC",           "ZINC"),
    (r"COBRE",          "SULFATO DE COBRE"),
    (r"MANGANESO",      "MANGANESO"),
    (r"HIERRO",         "HIERRO DEXTRANO"),
    (r"SELENIO",        "SELENITO DE SODIO"),
    (r"YODO",           "YODURO DE POTASIO"),
    (r"CROMO",          "PICOLINATO DE CROMO"),
    # Otros
    (r"BROMHEXIN",      "BROMHEXINA CLORHIDRATO"),
    (r"FUROSEMID",      "FUROSEMIDA"),
    (r"ATROPINA",       "ATROPINA SULFATO"),
    (r"EPINEFRINA|ADRENALINA", "EPINEFRINA"),
    (r"DIGOXINA",       "DIGOXINA"),
    (r"SILIMARINA",     "SILIMARINA"),
    (r"OMEPRAZOL",      "OMEPRAZOL"),
    (r"RANITIDINA",     "RANITIDINA"),
    (r"METOCLOPRAMID",  "METOCLOPRAMIDA"),
    (r"DIFENHIDRAMINA", "DIFENHIDRAMINA"),
    (r"CLORFENAMINA",   "CLORFENAMINA"),
    (r"DEXCLORFENIRAMINA", "DEXCLORFENIRAMINA"),
    (r"AZITROMICIN",    "AZITROMICINA"),
    (r"CEFAZOLINA",     "CEFAZOLINA"),
    (r"SULFATO DE ZINC","SULFATO DE ZINC"),
    (r"CLORURO DE SODIO","CLORURO DE SODIO"),
]
# Compilar los patrones de simulación una sola vez
SIMULACION_COMPILADA: list[tuple[re.Pattern, str]] = [
    (re.compile(patron, re.IGNORECASE), componente)
    for patron, componente in SIMULACION_DROGAS
]

# Alias de columnas → nombre canónico
ALIAS_COLUMNAS: dict[str, str] = {
    # PrincipioActivo
    "principio activo": "PrincipioActivo",
    "principioactivo": "PrincipioActivo",
    "principio": "PrincipioActivo",
    "producto": "PrincipioActivo",
    "nombre": "PrincipioActivo",
    # Presentacion
    "presentacion": "Presentacion",
    "presentación": "Presentacion",
    "forma farmaceutica": "Presentacion",
    "forma farmacéutica": "Presentacion",
    "forma": "Presentacion",
    # ComponenteMolecular
    "componente molecular": "ComponenteMolecular",
    "componentemolecular": "ComponenteMolecular",
    "composicion": "ComponenteMolecular",
    "composición": "ComponenteMolecular",
    "ingrediente activo": "ComponenteMolecular",
    "ingredienteactivo": "ComponenteMolecular",
    "ingrediente": "ComponenteMolecular",
    "principio activo molecular": "ComponenteMolecular",
}


# ---------------------------------------------------------------------------
# Funciones de normalización
# ---------------------------------------------------------------------------

def normalizar(texto: str) -> str:
    """Normalización básica: mayúsculas, sin tildes, sin porcentajes, espacios limpios."""
    if not isinstance(texto, str):
        return ""
    texto = texto.upper()
    texto = quitar_acentos(texto)
    texto = re.sub(r"\d+[.,]?\d*\s*%", "", texto)
    texto = re.sub(r"-\s*$", "", texto)
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_pa(texto: str) -> str:
    """Normalización química ampliada del PrincipioActivo:
    elimina prefijos estereoquímicos, sufijos de sales, calificadoras y prefijos de marca.
    """
    base = normalizar(texto)
    base = PREFIJOS_QUIRALES.sub("", base).strip()
    palabras = [EQUIVALENCIAS.get(p, p) for p in base.split()]
    base = " ".join(palabras)
    prev = None
    while prev != base:
        prev = base
        base = SUFIJOS_SALES.sub("", base).strip()
    base = PALABRAS_NO_MOLECULARES.sub("", base).strip()
    base = PREFIJOS_MARCA.sub("", base).strip()
    base = re.sub(r"\s+", " ", base).strip()
    base = re.sub(r"\s+\d+[\.,]?\d*\s*$", "", base).strip()
    return base


def simular_componente_molecular(principio_activo: str) -> str:
    """Genera un ComponenteMolecular coherente cuando no se puede obtener de los datos.

    Estrategia:
    1. Buscar el nombre del PA en el diccionario de drogas conocidas.
    2. Si no se identifica → "COMPUESTO GENERICO_" + nombre normalizado del PA.
    """
    pa_norm = normalizar(principio_activo)
    if not pa_norm:
        return "COMPUESTO GENERICO_DESCONOCIDO"

    for patron, componente in SIMULACION_COMPILADA:
        if patron.search(pa_norm):
            return componente

    # Fallback genérico
    return f"COMPUESTO GENERICO_{pa_norm}"


def extraer_principio_activo(presentacion: str) -> str:
    """Extrae el nombre del PA desde el string de presentación.
    Formato: 'NOMBRE DEL PRODUCTO - ENVASE - CANTIDAD'
    """
    if not isinstance(presentacion, str):
        return ""
    return presentacion.split(" - ", 1)[0].strip()


def detectar_y_renombrar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra columnas usando el diccionario de alias hacia el nombre canónico."""
    rename = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if col_lower in ALIAS_COLUMNAS:
            rename[col] = ALIAS_COLUMNAS[col_lower]
    return df.rename(columns=rename) if rename else df


# ---------------------------------------------------------------------------
# Lectura de archivos Reporte (fuente de ComponenteMolecular)
# ---------------------------------------------------------------------------

def _detectar_header_row(df_raw: pd.DataFrame, marcadores: list[str]) -> int | None:
    """Detecta la fila de encabezado buscando los marcadores dados."""
    for idx, row in df_raw.iterrows():
        vals_lower = [str(v).strip().lower() for v in row]
        if any(m.lower() in v for m in marcadores for v in vals_lower):
            return idx
    return None


def leer_reporte_archivo(ruta: str) -> pd.DataFrame:
    """Lee un archivo Reporte_Por_Principio_Activo y devuelve un DataFrame normalizado."""
    categoria = (
        os.path.basename(ruta)
        .replace("Reporte_Por_Principio_Activo_2024_", "")
        .replace(".xlsx", "")
    )
    try:
        xl = pd.ExcelFile(ruta)
    except Exception as exc:
        print(f"  [ERROR] No se pudo abrir {ruta}: {exc}")
        return pd.DataFrame()

    frames = []
    for hoja in xl.sheet_names:
        try:
            df_raw = pd.read_excel(ruta, sheet_name=hoja, header=None)
        except Exception as exc:
            print(f"  [ERROR] Hoja '{hoja}' en {ruta}: {exc}")
            continue

        header_row = _detectar_header_row(df_raw, ["Componente Molecular", "ComponenteMolecular"])
        if header_row is None:
            continue

        df = pd.read_excel(ruta, sheet_name=hoja, header=header_row)
        df = detectar_y_renombrar_columnas(df)

        # Renombrar columnas de ventas para evitar colisiones al unir
        rename_ventas = {}
        for col in df.columns:
            col_s = str(col).strip()
            if col_s in ("Primer Trimestre", "Segundo Trimestre", "Tercer Trimestre",
                         "Cuarto Trimestre", "Total Ventas", "Numero"):
                rename_ventas[col] = col_s.replace(" ", "") + "_Reporte"
        df = df.rename(columns=rename_ventas)

        if "ComponenteMolecular" not in df.columns:
            continue

        df = df[df["ComponenteMolecular"].notna()].copy()
        df["Categoria"] = categoria
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def cargar_todos_los_reportes(directorio: str) -> pd.DataFrame:
    """Carga todos los archivos Reporte_Por_Principio_Activo_*.xlsx."""
    patron = os.path.join(directorio, "Reporte_Por_Principio_Activo_2024_*.xlsx")
    archivos = sorted(glob.glob(patron))
    print(f"Se encontraron {len(archivos)} archivos de Reporte.")
    partes = []
    for ruta in archivos:
        cat = os.path.basename(ruta).replace("Reporte_Por_Principio_Activo_2024_", "").replace(".xlsx", "")
        print(f"  Leyendo: {cat}")
        df = leer_reporte_archivo(ruta)
        if not df.empty:
            partes.append(df)

    if not partes:
        return pd.DataFrame()

    resultado = pd.concat(partes, ignore_index=True)
    resultado = resultado[resultado["ComponenteMolecular"].astype(str).str.strip() != ""]
    resultado.drop_duplicates(subset=["Categoria", "ComponenteMolecular"], inplace=True)
    resultado.reset_index(drop=True, inplace=True)
    print(f"  → {len(resultado)} registros de ComponenteMolecular cargados.")
    return resultado


# ---------------------------------------------------------------------------
# Lectura de archivos del Mercado
# ---------------------------------------------------------------------------

def leer_mercado_archivo(ruta: str) -> pd.DataFrame:
    """Lee un archivo Excel del directorio Mercado (puede tener múltiples hojas)."""
    try:
        xl = pd.ExcelFile(ruta)
    except Exception as exc:
        print(f"  [ERROR] No se pudo abrir {ruta}: {exc}")
        return pd.DataFrame()

    frames = []
    for hoja in xl.sheet_names:
        try:
            df_raw = pd.read_excel(ruta, sheet_name=hoja, header=None)
        except Exception as exc:
            print(f"  [ERROR] Hoja '{hoja}' en {ruta}: {exc}")
            continue

        # Detectar fila de encabezado buscando columnas relevantes
        header_row = _detectar_header_row(
            df_raw,
            ["Presentacion", "Presentación", "PrincipioActivo", "Principio Activo",
             "ComponenteMolecular", "Componente Molecular"]
        )
        if header_row is None:
            print(f"  [AVISO] No se encontró encabezado reconocible en '{os.path.basename(ruta)}' hoja '{hoja}'")
            continue

        df = pd.read_excel(ruta, sheet_name=hoja, header=header_row)
        df = detectar_y_renombrar_columnas(df)

        # Sólo conservar filas con al menos una columna de interés con dato
        cols_utiles = [c for c in ("PrincipioActivo", "Presentacion", "ComponenteMolecular") if c in df.columns]
        if not cols_utiles:
            continue

        df = df.dropna(subset=cols_utiles, how="all")
        df["_ArchivoFuente"] = os.path.basename(ruta)
        df["_HojaFuente"] = hoja
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def cargar_directorio_mercado(directorio_mercado: str) -> pd.DataFrame:
    """Carga todos los Excel del directorio Mercado."""
    archivos = sorted(glob.glob(os.path.join(directorio_mercado, "*.xlsx")))
    print(f"Se encontraron {len(archivos)} archivo(s) en Mercado/.")
    partes = []
    for ruta in archivos:
        nombre = os.path.basename(ruta)
        print(f"  Leyendo: {nombre}")
        df = leer_mercado_archivo(ruta)
        if not df.empty:
            partes.append(df)

    if not partes:
        return pd.DataFrame()

    resultado = pd.concat(partes, ignore_index=True)
    print(f"  → {len(resultado)} filas cargadas del Mercado.")
    return resultado


# ---------------------------------------------------------------------------
# Motor de emparejamiento: PrincipioActivo ↔ ComponenteMolecular
# ---------------------------------------------------------------------------

def construir_indice_componentes(df_reportes: pd.DataFrame) -> list[dict]:
    """Construye índice de búsqueda sobre los ComponenteMolecular de los Reportes."""
    indice = []
    for _, fila in df_reportes.iterrows():
        cm = str(fila["ComponenteMolecular"]).strip()
        indice.append({
            "ComponenteMolecular": cm,
            "ComponenteMolecular_norm": normalizar(cm),
            "Categoria": fila.get("Categoria", ""),
        })
    return indice


def _es_simple(cm_norm: str) -> bool:
    return "," not in cm_norm


def buscar_componente(principio_activo: str, indice: list[dict]) -> str | None:
    """Busca el mejor ComponenteMolecular en el índice para un PrincipioActivo dado."""
    pa_basico = normalizar(principio_activo)
    pa_amplio = normalizar_pa(principio_activo)

    variantes = {v for v in (pa_basico, pa_amplio) if v}
    if not variantes:
        return None

    # 1. Coincidencia exacta
    for item in indice:
        if item["ComponenteMolecular_norm"] in variantes:
            return item["ComponenteMolecular"]

    # 2 y 3. Búsqueda por inclusión (más largo primero para mayor especificidad)
    for pa in sorted(variantes, key=len, reverse=True):
        if len(pa) < 4:
            continue
        simples = [it for it in indice if _es_simple(it["ComponenteMolecular_norm"]) and pa in it["ComponenteMolecular_norm"]]
        compuestos = [it for it in indice if not _es_simple(it["ComponenteMolecular_norm"]) and pa in it["ComponenteMolecular_norm"]]
        if simples:
            return min(simples, key=lambda x: len(x["ComponenteMolecular_norm"]))["ComponenteMolecular"]
        if compuestos:
            return min(compuestos, key=lambda x: len(x["ComponenteMolecular_norm"]))["ComponenteMolecular"]

    # 4. Palabras clave (≥6 chars) del PA en componentes simples
    for pa in sorted(variantes, key=len, reverse=True):
        palabras = [p for p in pa.split() if len(p) >= 6]
        if not palabras:
            continue
        mejores = [
            it for it in indice
            if _es_simple(it["ComponenteMolecular_norm"])
            and all(p in it["ComponenteMolecular_norm"] for p in palabras)
        ]
        if mejores:
            return min(mejores, key=lambda x: len(x["ComponenteMolecular_norm"]))["ComponenteMolecular"]

    return None


# ---------------------------------------------------------------------------
# Consolidación principal
# ---------------------------------------------------------------------------

def consolidar(directorio: str, directorio_mercado: str, ruta_salida: str) -> None:
    print("\n=== [1/4] Cargando archivos de Reporte (ComponenteMolecular) ===")
    df_reportes = cargar_todos_los_reportes(directorio)

    print("\n=== [2/4] Cargando archivos del Mercado (Presentaciones) ===")
    df = cargar_directorio_mercado(directorio_mercado)

    if df.empty:
        print("[ERROR] No se pudo cargar ningún archivo del Mercado. Abortando.")
        return

    # -----------------------------------------------------------------------
    # Extraer PrincipioActivo si no existe como columna explícita
    # -----------------------------------------------------------------------
    if "PrincipioActivo" not in df.columns:
        if "Presentacion" in df.columns:
            df["PrincipioActivo"] = df["Presentacion"].apply(
                lambda x: extraer_principio_activo(str(x)) if pd.notna(x) else ""
            )
        else:
            df["PrincipioActivo"] = ""

    # Normalizar texto en columnas clave
    for col in ("PrincipioActivo", "Presentacion"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
            df[col] = df[col].replace("NAN", "")

    # -----------------------------------------------------------------------
    # [3/4] Enriquecer con ComponenteMolecular desde los Reportes
    # -----------------------------------------------------------------------
    print("\n=== [3/4] Asociando ComponenteMolecular a cada PrincipioActivo ===")

    indice = construir_indice_componentes(df_reportes) if not df_reportes.empty else []

    # Inicializar ComponenteMolecularFinal con lo que ya exista en los datos
    if "ComponenteMolecular" in df.columns:
        df["ComponenteMolecularFinal"] = df["ComponenteMolecular"].astype(str).str.strip()
        df["ComponenteMolecularFinal"] = df["ComponenteMolecularFinal"].replace({"NAN": "", "nan": ""})
    else:
        df["ComponenteMolecularFinal"] = ""

    # Buscar en los Reportes para filas sin ComponenteMolecular
    busquedas_exitosas = 0
    for idx, fila in df.iterrows():
        if df.at[idx, "ComponenteMolecularFinal"]:
            continue
        pa = str(fila.get("PrincipioActivo", "") or "").strip()
        if not pa:
            continue
        cm = buscar_componente(pa, indice)
        if cm:
            df.at[idx, "ComponenteMolecularFinal"] = cm
            busquedas_exitosas += 1

    print(f"  → {busquedas_exitosas} filas enriquecidas desde Reportes.")

    # -----------------------------------------------------------------------
    # Propagación: si un PA ya tiene CM en alguna fila, propagar a las demás
    # -----------------------------------------------------------------------
    print("\n=== [3b] Propagación de ComponenteMolecular entre presentaciones ===")
    propagaciones = 0

    # Construir mapa PA → mejor CM (el más corto como proxy del más específico de molécula simple)
    pa_a_cm: dict[str, str] = {}
    for idx, fila in df.iterrows():
        pa = str(fila.get("PrincipioActivo", "") or "").strip()
        cm = str(fila.get("ComponenteMolecularFinal", "") or "").strip()
        if pa and cm:
            # Preferir CMs ya asignados (no genéricos)
            if pa not in pa_a_cm or (
                not pa_a_cm[pa].startswith("COMPUESTO GENERICO_")
                and len(cm) < len(pa_a_cm[pa])
            ):
                pa_a_cm[pa] = cm

    # Propagar
    for idx, fila in df.iterrows():
        pa = str(fila.get("PrincipioActivo", "") or "").strip()
        if not df.at[idx, "ComponenteMolecularFinal"] and pa in pa_a_cm:
            df.at[idx, "ComponenteMolecularFinal"] = pa_a_cm[pa]
            propagaciones += 1

    print(f"  → {propagaciones} filas completadas por propagación.")

    # -----------------------------------------------------------------------
    # Simulación: generar CM para los PAs que aún no tienen ninguno
    # -----------------------------------------------------------------------
    print("\n=== [3c] Simulación química para PAs sin ComponenteMolecular ===")
    simulaciones = 0

    for idx, fila in df.iterrows():
        if df.at[idx, "ComponenteMolecularFinal"]:
            continue
        pa = str(fila.get("PrincipioActivo", "") or "").strip()
        presentacion = str(fila.get("Presentacion", "") or "").strip()
        nombre_base = pa if pa else presentacion
        cm_simulado = simular_componente_molecular(nombre_base)
        df.at[idx, "ComponenteMolecularFinal"] = cm_simulado
        simulaciones += 1

    print(f"  → {simulaciones} componentes simulados.")

    # -----------------------------------------------------------------------
    # Limpieza final
    # -----------------------------------------------------------------------
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    # Eliminar duplicados: misma PA + Presentacion (si existen ambas columnas)
    dup_cols = [c for c in ("PrincipioActivo", "Presentacion") if c in df.columns]
    if dup_cols:
        df.drop_duplicates(subset=dup_cols, inplace=True)

    df.reset_index(drop=True, inplace=True)

    # Garantía final: ComponenteMolecularFinal nunca vacío
    mask_vacio = df["ComponenteMolecularFinal"].astype(str).str.strip() == ""
    if mask_vacio.any():
        for idx in df[mask_vacio].index:
            pa = str(df.at[idx, "PrincipioActivo"] if "PrincipioActivo" in df.columns else "")
            df.at[idx, "ComponenteMolecularFinal"] = simular_componente_molecular(pa)

    # Reordenar columnas: primero las canónicas, luego el resto
    col_prio = ["PrincipioActivo", "Presentacion", "ComponenteMolecularFinal"]
    otras = [c for c in df.columns if c not in col_prio]
    df = df[[c for c in col_prio if c in df.columns] + otras]

    # -----------------------------------------------------------------------
    # Guardar el resultado
    # -----------------------------------------------------------------------
    print(f"\n=== [4/4] Guardando resultado en: {ruta_salida} ===")
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Consolidado", index=False)

        if not df_reportes.empty:
            df_reportes.to_excel(writer, sheet_name="ReportesComponentes", index=False)

    total = len(df)
    con_cm = (df["ComponenteMolecularFinal"].astype(str).str.strip() != "").sum()
    print(f"  → Archivo guardado: {ruta_salida}")
    print(f"  → Filas en 'Consolidado': {total} | Con ComponenteMolecularFinal: {con_cm}/{total}")
    if not df_reportes.empty:
        print(f"  → Filas en 'ReportesComponentes': {len(df_reportes)}")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    consolidar(
        directorio=BASE_DIR,
        directorio_mercado=os.path.join(BASE_DIR, "Mercado"),
        ruta_salida=os.path.join(BASE_DIR, "PrincipioActivoAnalisisMercado2026.xlsx"),
    )
