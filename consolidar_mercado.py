"""
Consolidador de Principios Activos del Mercado 2026
=====================================================
Lee el análisis de mercado (Mercado/AnalisisDelMercadoSector.xlsx) y todos los
archivos de principios activos (Reporte_Por_Principio_Activo_2024_*.xlsx).

Genera PrincipioActivoAnalisisMercado2026.xlsx conservando TODAS las columnas
originales del análisis de mercado y agregando únicamente la columna
ComponenteMolecular, obtenida cruzando el nombre de la presentación contra los
componentes moleculares conocidos de los archivos Reporte.
"""

import os
import re
import glob
import logging
import unicodedata

import pandas as pd

# ---------------------------------------------------------------------------
# Configuración de logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MERCADO_DIR = os.path.join(SCRIPT_DIR, "Mercado")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "PrincipioActivoAnalisisMercado2026.xlsx")


# ---------------------------------------------------------------------------
# Utilidades de normalización
# ---------------------------------------------------------------------------

def _normalizar(texto: str) -> str:
    """Quita tildes, convierte a mayúsculas y colapsa espacios."""
    if not isinstance(texto, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", sin_tildes).strip().upper()


def _extraer_nombre_producto(presentacion: str) -> str:
    """Extrae el identificador del producto desde el campo Presentacion.

    Ejemplo:
        'DL-METHIONINA 99% - FUNDA - 1 KG'  →  'DL-METHIONINA'
        'IVERMECTINA 1% - FRASCO - 500 ML'   →  'IVERMECTINA'
        'CEVAC VITABRON L - VIAL - 200 ML'   →  'CEVAC VITABRON L'
    """
    if not isinstance(presentacion, str):
        return ""
    # Tomar la parte antes del primer ' - '
    nombre = presentacion.split(" - ")[0].strip()
    # Eliminar indicadores de concentración/pureza al final (p.ej. "99%", "60%", "1%")
    nombre = re.sub(r"\s+\d+[\.,]?\d*\s*%.*$", "", nombre).strip()
    return _normalizar(nombre)


# ---------------------------------------------------------------------------
# Construcción del índice de ComponenteMolecular desde archivos Reporte
# ---------------------------------------------------------------------------

def construir_indice_componentes(patron_reporte: str) -> tuple[dict, set]:
    """Lee todos los archivos Reporte y construye un índice de búsqueda.

    Devuelve una tupla (indice, moleculas):
    - indice: dict  componente_normalizado → ComponenteMolecular_completo
    - moleculas: set de moléculas individuales (tokens separados por coma)
    """
    componentes_set: set[str] = set()
    archivos = sorted(glob.glob(patron_reporte))
    logger.info("Leyendo componentes desde %d archivos Reporte …", len(archivos))

    for ruta in archivos:
        try:
            excel = pd.ExcelFile(ruta, engine="openpyxl")
        except Exception as exc:
            logger.error("No se pudo abrir '%s': %s", os.path.basename(ruta), exc)
            continue

        for sheet_name in excel.sheet_names:
            try:
                raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=str)
                header_idx = None
                for idx, row in raw.iterrows():
                    if str(row.iloc[0]).strip().lower() == "numero":
                        header_idx = idx
                        break
                if header_idx is None:
                    continue
                df = pd.read_excel(excel, sheet_name=sheet_name, header=header_idx, dtype=str)
                col_comp = next(
                    (c for c in df.columns if "componente" in c.lower()),
                    None,
                )
                if col_comp is None:
                    continue
                for valor in df[col_comp].dropna():
                    norma = _normalizar(valor)
                    if norma:
                        componentes_set.add(norma)
            except Exception as exc:
                logger.error(
                    "Error en hoja '%s' de '%s': %s",
                    sheet_name,
                    os.path.basename(ruta),
                    exc,
                )

    logger.info("  Componentes únicos cargados: %d", len(componentes_set))

    # Índice directo: ComponenteMolecular completo → sí mismo
    indice: dict[str, str] = {comp: comp for comp in componentes_set}

    # Moléculas individuales (tokens): cada sub-componente separado por coma
    moleculas: set[str] = set()
    for comp in componentes_set:
        for parte in comp.split(","):
            parte_clean = re.sub(r"\s+", " ", parte.strip())
            if len(parte_clean) > 3:
                moleculas.add(parte_clean)

    return indice, moleculas


def buscar_componente(nombre_producto: str, indice: dict, moleculas: set) -> str:
    """Intenta encontrar el ComponenteMolecular para un nombre de producto dado.

    Estrategia (en orden de prioridad):
    1. Coincidencia exacta del nombre del producto con el ComponenteMolecular completo.
    2. El nombre del producto coincide exactamente con un sub-componente individual
       (token separado por coma) de algún ComponenteMolecular.
    3. El nombre del producto es prefijo de algún ComponenteMolecular completo
       (límite de palabra: espacio, coma o fin de cadena).
    Si no hay coincidencia, devuelve el propio nombre del producto (mejor esfuerzo).
    """
    if not nombre_producto:
        return ""

    # 1. Coincidencia exacta con ComponenteMolecular completo
    if nombre_producto in indice:
        return indice[nombre_producto]

    # 2. El nombre coincide exactamente con una molécula individual (token exacto)
    if nombre_producto in moleculas:
        # Buscar el ComponenteMolecular completo que contiene esta molécula como único token
        for comp in sorted(indice, key=len):
            tokens = {t.strip() for t in comp.split(",")}
            if nombre_producto in tokens and len(tokens) == 1:
                return comp
        # Es parte de una mezcla: devolver el nombre del producto como está
        return nombre_producto

    # 3. El nombre del producto es prefijo exacto de algún ComponenteMolecular
    #    (protegido con límite de palabra: espacio, coma o fin de cadena)
    patron = re.compile(
        r"^" + re.escape(nombre_producto) + r"(?:\s|,|$)",
        re.IGNORECASE,
    )
    for comp in sorted(indice, key=len):
        if patron.match(comp):
            return indice[comp]

    # Sin coincidencia: usar el propio nombre del producto como ComponenteMolecular
    return nombre_producto


# ---------------------------------------------------------------------------
# Lectura del análisis de mercado conservando columnas originales
# ---------------------------------------------------------------------------

def leer_analisis_mercado(ruta_archivo: str) -> pd.DataFrame:
    """Lee un archivo de análisis de mercado y devuelve un DataFrame con
    TODAS sus columnas originales, sin renombrar ni transformar."""
    frames = []

    try:
        excel = pd.ExcelFile(ruta_archivo, engine="openpyxl")
    except Exception as exc:
        logger.error("No se pudo abrir '%s': %s", ruta_archivo, exc)
        return pd.DataFrame()

    for sheet_name in excel.sheet_names:
        try:
            raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=str)

            # Detectar fila de encabezado buscando "Rank" o "Presentacion"
            header_idx = None
            for idx, row in raw.iterrows():
                row_lower = [str(c).strip().lower() for c in row if pd.notna(c)]
                if "rank" in row_lower or "presentacion" in row_lower:
                    header_idx = idx
                    break

            if header_idx is None:
                logger.warning(
                    "Hoja '%s' en '%s': encabezado no encontrado, se omite.",
                    sheet_name,
                    os.path.basename(ruta_archivo),
                )
                continue

            # Leer con los encabezados originales del archivo
            df = pd.read_excel(excel, sheet_name=sheet_name, header=header_idx, dtype=str)

            # Eliminar fila de totales y filas vacías en Presentacion
            col_pres = next(
                (c for c in df.columns if str(c).strip().lower() == "presentacion"),
                None,
            )
            if col_pres:
                df = df[df[col_pres].notna()].copy()
                df = df[
                    ~df[col_pres].str.strip().str.upper().isin(["TOTALES", ""])
                ].copy()
                df[col_pres] = df[col_pres].str.strip()

            # Eliminar columnas completamente vacías que pandas puede generar
            df.dropna(axis=1, how="all", inplace=True)

            frames.append(df)

        except Exception as exc:
            logger.error(
                "Error leyendo hoja '%s' en '%s': %s",
                sheet_name,
                os.path.basename(ruta_archivo),
                exc,
            )

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def consolidar():
    logger.info("=== Consolidador de Principios Activos del Mercado 2026 ===")

    # 1. Construir índice de ComponenteMolecular desde archivos Reporte
    patron_reporte = os.path.join(SCRIPT_DIR, "Reporte_Por_Principio_Activo_2024_*.xlsx")
    indice_componentes, moleculas = construir_indice_componentes(patron_reporte)

    # 2. Leer todos los archivos de análisis de mercado (carpeta Mercado)
    archivos_mercado = sorted(glob.glob(os.path.join(MERCADO_DIR, "*.xlsx")))
    logger.info("Archivos en carpeta Mercado: %d", len(archivos_mercado))

    dfs = []
    for ruta in archivos_mercado:
        logger.info("  Leyendo: %s", os.path.basename(ruta))
        df = leer_analisis_mercado(ruta)
        if not df.empty:
            dfs.append(df)

    if not dfs:
        logger.error("No se encontró ningún dato en la carpeta Mercado. Saliendo.")
        return

    df_mercado = pd.concat(dfs, ignore_index=True)
    logger.info("Total filas del análisis de mercado: %d", len(df_mercado))

    # 3. Agregar columna ComponenteMolecular
    col_pres = next(
        (c for c in df_mercado.columns if str(c).strip().lower() == "presentacion"),
        None,
    )
    if col_pres is None:
        logger.error("No se encontró la columna 'Presentacion' en los datos de mercado.")
        return

    logger.info("Asociando ComponenteMolecular a cada Presentacion …")

    def _asignar_componente(presentacion: str) -> str:
        nombre_prod = _extraer_nombre_producto(presentacion)
        return buscar_componente(nombre_prod, indice_componentes, moleculas)

    df_mercado["ComponenteMolecular"] = df_mercado[col_pres].apply(_asignar_componente)

    # Mover ComponenteMolecular justo después de Presentacion
    cols = list(df_mercado.columns)
    cols.remove("ComponenteMolecular")
    pos = cols.index(col_pres) + 1
    cols.insert(pos, "ComponenteMolecular")
    df_mercado = df_mercado[cols]

    # Eliminar duplicados
    antes = len(df_mercado)
    df_mercado.drop_duplicates(subset=[col_pres], inplace=True)
    logger.info("Filas únicas (por Presentacion): %d → %d", antes, len(df_mercado))

    # 4. Guardar el archivo de salida
    logger.info("Generando '%s' …", OUTPUT_FILE)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_mercado.to_excel(writer, sheet_name="AnalisisMercado2026", index=False)
        logger.info(
            "  Hoja 'AnalisisMercado2026': %d filas, %d columnas",
            len(df_mercado),
            len(df_mercado.columns),
        )

    logger.info("Archivo generado exitosamente: %s", OUTPUT_FILE)


if __name__ == "__main__":
    consolidar()
