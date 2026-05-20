"""
Genera PrincipioActivoAnalisisMercado2026.xlsx consolidando:
- Todos los archivos Reporte_Por_Principio_Activo_2024_*.xlsx (ComponenteMolecular + ventas por categoria)
- Mercado/AnalisisDelMercadoSector.xlsx (Presentaciones + datos de mercado)

Para cada Presentacion en el archivo de Mercado se extrae el PrincipioActivo
(parte antes del primer ' - ') y se busca el ComponenteMolecular correspondiente
en los archivos de Reporte, normalizando los nombres para comparacion.
"""

import glob
import os
import re
import unicodedata

import pandas as pd


# ---------------------------------------------------------------------------
# Utilidades de normalización
# ---------------------------------------------------------------------------

def quitar_acentos(texto: str) -> str:
    """Elimina tildes/diacríticos manteniendo la letra base."""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# Equivalencias ortográficas: español ↔ inglés / variantes de escritura
# clave = forma a normalizar → valor = forma canónica en los Reportes
EQUIVALENCIAS = {
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
    "FLORFENICOL": "FLORFENICOL",
    "TILMICOSIN": "TILMICOSINA",
    "TYLOSIN": "TILOSINA",
    "LINCOMYCIN": "LINCOMICINA",
    "SPECTINOMYCIN": "ESPECTINOMICINA",
    "COLISTIN": "COLISTINA",
    "DEXAMETHASONE": "DEXAMETASONA",
    "PREDNISOLONE": "PREDNISOLONA",
    "MELOXICAM": "MELOXICAM",
    "KETAMINE": "KETAMINA",
    "XYLAZINE": "XILAZINA",
    "BUTORPHANOL": "BUTORFANOL",
    "OXYTOCIN": "OXITOCINA",
    "PROGESTERONE": "PROGESTERONA",
    "TESTOSTERONE": "TESTOSTERONA",
    "ESTRADIOL": "ESTRADIOL",
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
    "SELENIUM": "SELENIO",
    "IODINE": "YODO",
}

# Palabras calificadoras (especie, vía, marca) que no forman parte del nombre molecular
PALABRAS_NO_MOLECULARES = re.compile(
    r"\b(AVICOLA|AVIAR|BOVINO|BOVINA|CANINO|CANINA|FELINO|FELINA|EQUINO|EQUINA|"
    r"PORCINO|PORCINA|OVINO|OVINA|CAPRINO|CAPRINA|CUNICULAR|"
    r"INYECTABLE|ORAL|TOPICO|TOPICA|INTRAVENOSO|SUBCUTANEO|"
    r"FORTE|PLUS|ULTRA|SUPER|MAX|PRO|PREMIUM|EXTRA|SPECIAL|"
    r"AVICOLA|PECUARIO|VETERINARIO|VETERINARIA)\b",
    re.IGNORECASE,
)

# Prefijos de marca conocidos (cortos, al inicio del texto)
PREFIJOS_MARCA = re.compile(r"^(LH|VM|BM|AG|QSI|MC)\s*", re.IGNORECASE)

# Prefijos estereoquímicos/de forma que no definen la molécula
PREFIJOS_QUIRALES = re.compile(
    r"^(DL|L|D|R|S|MESO|CIS|TRANS|N|O|S|ALFA|BETA|GAMMA|DELTA)-\s*", re.IGNORECASE
)

# Sufijos de sales / formas farmacéuticas que no cambian el PA
SUFIJOS_SALES = re.compile(
    r"\s+(HCL|CLORHIDRATO|HIDROCLORURO|SULFATO|FOSFATO|TARTRATO|CITRATO|LACTATO|"
    r"GLUCONATO|ACETATO|NITRATO|BROMURO|YODURO|MALEATO|FUMARATO|SUCCINATO|"
    r"SÓDICO|SODICO|POTÁSICO|POTASICO|MAGNÉSICO|MAGNESICO|AMÓNICO|AMONICO|"
    r"TRIHIDRATADO|TRIHIDRATADA|MONOHIDRATADO|HIDRATADO|ANHIDRO|BASE|"
    r"MICRONIZADO|MICRONIZADA|GRANULADO|GRANULADA|POLVO)$",
    re.IGNORECASE,
)


def normalizar(texto: str) -> str:
    """Normaliza un nombre para comparación:
    - Mayúsculas
    - Sin tildes
    - Sin porcentajes ni concentraciones (p.ej. '99%', '60%')
    - Espacios únicos y sin bordes
    """
    if not isinstance(texto, str):
        return ""
    texto = texto.upper()
    texto = quitar_acentos(texto)
    # Eliminar porcentajes y concentraciones numéricas (98%, 98,5%, etc.)
    texto = re.sub(r"\d+[.,]?\d*\s*%", "", texto)
    # Eliminar guiones y hifens sobrantes al final de palabra
    texto = re.sub(r"-\s*$", "", texto)
    # Eliminar espacios múltiples
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_pa(texto: str) -> str:
    """Normalización ampliada para el PrincipioActivo extraído de la Presentacion.
    Aplica la normalización básica + equivalencias ortográficas + limpieza de
    prefijos estereoquímicos y sufijos de sales.
    """
    base = normalizar(texto)

    # 1. Remover prefijos estereoquímicos de una sola letra al inicio (L-, DL-, etc.)
    #    Se hace ANTES de las equivalencias para que "DL-METHIONINA" → "METHIONINA"
    base = PREFIJOS_QUIRALES.sub("", base).strip()

    # 2. Aplicar equivalencias ortográficas (palabra a palabra)
    palabras = base.split()
    palabras = [EQUIVALENCIAS.get(p, p) for p in palabras]
    base = " ".join(palabras)

    # 3. Remover sufijos de sales / formas farmacéuticas
    prev = None
    while prev != base:
        prev = base
        base = SUFIJOS_SALES.sub("", base).strip()

    # 4. Remover palabras calificadoras (especie, vía de administración, marca)
    base = PALABRAS_NO_MOLECULARES.sub("", base).strip()
    # 5. Remover prefijos de marca conocidos al inicio
    base = PREFIJOS_MARCA.sub("", base).strip()
    # 6. Limpiar espacios múltiples tras los reemplazos
    base = re.sub(r"\s+", " ", base).strip()

    # 7. Limpiar texto numérico de concentración al final (ej. "LISINA 98")
    base = re.sub(r"\s+\d+[\.,]?\d*\s*$", "", base).strip()

    return base


def extraer_principio_activo(presentacion: str) -> str:
    """Extrae el nombre del principio activo desde el string de presentación.
    Formato esperado: 'NOMBRE DEL PRODUCTO - ENVASE - CANTIDAD'
    """
    if not isinstance(presentacion, str):
        return ""
    # Tomar la parte antes del primer ' - '
    partes = presentacion.split(" - ", 1)
    return partes[0].strip()


# ---------------------------------------------------------------------------
# Lectura de archivos de Reporte (ComponenteMolecular por Categoria)
# ---------------------------------------------------------------------------

def leer_reporte_archivo(ruta: str) -> pd.DataFrame:
    """Lee un archivo Reporte_Por_Principio_Activo y devuelve un DataFrame
    con columnas: Categoria, ComponenteMolecular, PrimerTrimestre,
    SegundoTrimestre, TercerTrimestre, CuartoTrimestre, TotalVentas
    """
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

        # Detectar fila de encabezado buscando la columna 'Componente Molecular'
        header_row = None
        for idx, row in df_raw.iterrows():
            vals = [str(v).strip() for v in row]
            if any("Componente Molecular" in v for v in vals):
                header_row = idx
                break

        if header_row is None:
            print(f"  [AVISO] No se encontró encabezado en {ruta} hoja '{hoja}'")
            continue

        df = pd.read_excel(ruta, sheet_name=hoja, header=header_row)

        # Renombrar columnas clave
        rename = {}
        for col in df.columns:
            col_s = str(col).strip()
            if col_s == "Componente Molecular":
                rename[col] = "ComponenteMolecular"
            elif col_s == "Primer Trimestre":
                rename[col] = "PrimerTrimestre_Reporte"
            elif col_s == "Segundo Trimestre":
                rename[col] = "SegundoTrimestre_Reporte"
            elif col_s == "Tercer Trimestre":
                rename[col] = "TercerTrimestre_Reporte"
            elif col_s == "Cuarto Trimestre":
                rename[col] = "CuartoTrimestre_Reporte"
            elif col_s == "Total Ventas":
                rename[col] = "TotalVentas_Reporte"
            elif col_s == "Numero":
                rename[col] = "Numero_Reporte"
        df = df.rename(columns=rename)

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
    # Eliminar filas sin ComponenteMolecular y duplicados exactos
    resultado = resultado[resultado["ComponenteMolecular"].astype(str).str.strip() != ""]
    resultado.drop_duplicates(subset=["Categoria", "ComponenteMolecular"], inplace=True)
    resultado.reset_index(drop=True, inplace=True)
    print(f"  → {len(resultado)} filas de ComponenteMolecular cargadas.")
    return resultado


# ---------------------------------------------------------------------------
# Lectura del archivo de Mercado
# ---------------------------------------------------------------------------

def cargar_mercado(ruta: str) -> pd.DataFrame:
    """Carga el archivo de mercado y estandariza columnas."""
    try:
        df_raw = pd.read_excel(ruta, header=None)
    except Exception as exc:
        print(f"[ERROR] No se pudo abrir {ruta}: {exc}")
        return pd.DataFrame()

    # Detectar fila de encabezado buscando 'Presentacion'
    header_row = None
    for idx, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row]
        if any("Presentacion" in v for v in vals):
            header_row = idx
            break

    if header_row is None:
        print(f"[AVISO] No se encontró encabezado 'Presentacion' en {ruta}")
        return pd.DataFrame()

    df = pd.read_excel(ruta, header=header_row)
    df = df[df["Presentacion"].notna()].copy()
    df.reset_index(drop=True, inplace=True)
    print(f"  → {len(df)} presentaciones cargadas del Mercado.")
    return df


# ---------------------------------------------------------------------------
# Motor de emparejamiento: Presentacion ↔ ComponenteMolecular
# ---------------------------------------------------------------------------

def construir_indice_componentes(df_reportes: pd.DataFrame) -> list[dict]:
    """Construye una lista de dicts con la clave normalizada de cada componente."""
    indice = []
    for _, fila in df_reportes.iterrows():
        cm = str(fila["ComponenteMolecular"]).strip()
        indice.append(
            {
                "ComponenteMolecular": cm,
                "ComponenteMolecular_norm": normalizar(cm),
                "Categoria": fila.get("Categoria", ""),
                "_fila": fila,
            }
        )
    return indice


def _es_componente_simple(cm_norm: str) -> bool:
    """Devuelve True si el ComponenteMolecular normalizado es de molécula única
    (no contiene comas, es decir, no es una combinación de varios activos).
    """
    return "," not in cm_norm


def buscar_mejor_componente(principio_activo: str, indice: list[dict]) -> dict | None:
    """Busca en el índice el ComponenteMolecular que mejor corresponde al
    principio activo extraído de la Presentacion.

    Estrategia (en orden de prioridad):
    1. Coincidencia exacta del PA normalizado con el CM normalizado
       (tanto con normalización simple como con normalización ampliada)
    2. El CM es de molécula única y el PA normalizado está contenido en él
       (mínimo 4 chars, para evitar coincidencias espurias)
    3. El PA normalizado es substring del CM (moléculas compuestas),
       priorizando el CM más corto (más específico)
    4. Búsqueda por palabras significativas del PA (≥6 chars) en CM únicos
    """
    # Generar variantes normalizadas del PA
    pa_basico = normalizar(principio_activo)
    pa_amplio = normalizar_pa(principio_activo)

    variantes = {pa_basico, pa_amplio}
    variantes.discard("")

    if not variantes:
        return None

    # --- 1. Coincidencia exacta (cualquier variante) ---
    for item in indice:
        cm_n = item["ComponenteMolecular_norm"]
        if cm_n in variantes:
            return item

    # --- 2 y 3. Búsqueda por inclusión/substring ---
    for pa_norm in sorted(variantes, key=len, reverse=True):  # más específica primero
        if len(pa_norm) < 4:
            continue

        simples = [item for item in indice if _es_componente_simple(item["ComponenteMolecular_norm"]) and pa_norm in item["ComponenteMolecular_norm"]]
        compuestos = [item for item in indice if not _es_componente_simple(item["ComponenteMolecular_norm"]) and pa_norm in item["ComponenteMolecular_norm"]]

        # Preferir componentes simples que contengan el PA
        if simples:
            return min(simples, key=lambda x: len(x["ComponenteMolecular_norm"]))
        if compuestos:
            return min(compuestos, key=lambda x: len(x["ComponenteMolecular_norm"]))

    # --- 4. Palabras significativas (≥6 chars) del PA en CMs simples ---
    for pa_norm in sorted(variantes, key=len, reverse=True):
        palabras = [p for p in pa_norm.split() if len(p) >= 6]
        if not palabras:
            continue
        mejores = [
            item for item in indice
            if _es_componente_simple(item["ComponenteMolecular_norm"])
            and all(p in item["ComponenteMolecular_norm"] for p in palabras)
        ]
        if mejores:
            return min(mejores, key=lambda x: len(x["ComponenteMolecular_norm"]))

    return None


# ---------------------------------------------------------------------------
# Consolidación principal
# ---------------------------------------------------------------------------

def consolidar(directorio: str, ruta_mercado: str, ruta_salida: str) -> None:
    print("\n=== Cargando archivos de Reporte (ComponenteMolecular) ===")
    df_reportes = cargar_todos_los_reportes(directorio)

    print("\n=== Cargando archivo de Mercado (Presentaciones) ===")
    df_mercado = cargar_mercado(ruta_mercado)

    if df_mercado.empty:
        print("[ERROR] No se pudo cargar el archivo de Mercado. Abortando.")
        return

    # Construir índice de componentes para búsqueda
    indice = construir_indice_componentes(df_reportes) if not df_reportes.empty else []

    # -----------------------------------------------------------------------
    # Enriquecer cada fila del Mercado con PrincipioActivo + ComponenteMolecular
    # -----------------------------------------------------------------------
    print("\n=== Asociando ComponenteMolecular a cada Presentacion ===")
    principios = []
    componentes = []
    categorias = []
    coincidencias = 0

    for presentacion in df_mercado["Presentacion"]:
        pa = extraer_principio_activo(str(presentacion))
        match = buscar_mejor_componente(pa, indice)
        principios.append(pa)
        if match:
            componentes.append(match["ComponenteMolecular"])
            categorias.append(match["Categoria"])
            coincidencias += 1
        else:
            componentes.append("NO IDENTIFICADO")
            categorias.append("")

    df_mercado.insert(0, "PrincipioActivo", principios)
    df_mercado.insert(1, "ComponenteMolecular", componentes)
    df_mercado.insert(2, "Categoria", categorias)

    total = len(df_mercado)
    print(f"  → {coincidencias}/{total} presentaciones con ComponenteMolecular identificado.")

    # -----------------------------------------------------------------------
    # Limpiar y estandarizar el DataFrame resultante
    # -----------------------------------------------------------------------
    # Estandarizar nombres de columnas: sin espacios en extremos
    df_mercado.columns = [str(c).strip() for c in df_mercado.columns]

    # Eliminar filas completamente vacías
    df_mercado.dropna(how="all", inplace=True)

    # Eliminar duplicados en combinación PrincipioActivo + Presentacion
    df_mercado.drop_duplicates(subset=["PrincipioActivo", "Presentacion"], inplace=True)
    df_mercado.reset_index(drop=True, inplace=True)

    # -----------------------------------------------------------------------
    # Guardar el resultado
    # -----------------------------------------------------------------------
    print(f"\n=== Guardando resultado en: {ruta_salida} ===")
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_mercado.to_excel(writer, sheet_name="Consolidado", index=False)

        # Hoja auxiliar: todos los ComponenteMolecular de los Reportes
        if not df_reportes.empty:
            df_reportes.to_excel(writer, sheet_name="ReportesComponentes", index=False)

    print(f"  → Archivo guardado: {ruta_salida}")
    print(f"  → Filas en hoja 'Consolidado': {len(df_mercado)}")
    if not df_reportes.empty:
        print(f"  → Filas en hoja 'ReportesComponentes': {len(df_reportes)}")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    consolidar(
        directorio=BASE_DIR,
        ruta_mercado=os.path.join(BASE_DIR, "Mercado", "AnalisisDelMercadoSector.xlsx"),
        ruta_salida=os.path.join(BASE_DIR, "PrincipioActivoAnalisisMercado2026.xlsx"),
    )
