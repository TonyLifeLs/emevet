"""
Consolidador de Principios Activos del Mercado 2026
=====================================================
Lee todos los archivos Excel de principios activos (Reporte_Por_Principio_Activo_2024_*.xlsx)
y el análisis de mercado por sector (Mercado/AnalisisDelMercadoSector.xlsx),
unifica la información y genera PrincipioActivoAnalisisMercado2026.xlsx.
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

# Número de filas de encabezado institucional antes de la fila de columnas
# en los archivos Reporte_Por_Principio_Activo_2024_*.xlsx
HEADER_ROWS_REPORTE = 4


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def normalizar_texto(texto: str) -> str:
    """Quita tildes, convierte a mayúsculas y elimina espacios redundantes."""
    if not isinstance(texto, str):
        return texto
    nfkd = unicodedata.normalize("NFKD", texto)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", sin_tildes).strip().upper()


def normalizar_componente(componente: str) -> str:
    """Normaliza el nombre de un componente molecular."""
    return normalizar_texto(componente)


def extraer_principio_activo(nombre_archivo: str) -> str:
    """Extrae el nombre del principio activo a partir del nombre de archivo.

    Ejemplo:
        Reporte_Por_Principio_Activo_2024_ANTIPARASITARIOS.xlsx
        → ANTIPARASITARIOS
    """
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    # Patrón: Reporte_Por_Principio_Activo_YYYY_<CATEGORIA>
    match = re.search(r"Reporte_Por_Principio_Activo_\d{4}_(.+)$", base, re.IGNORECASE)
    if match:
        categoria = match.group(1).replace("_", " ")
        return normalizar_texto(categoria)
    return normalizar_texto(base)


# ---------------------------------------------------------------------------
# Lectura de archivos Reporte
# ---------------------------------------------------------------------------

def leer_reporte(ruta_archivo: str) -> pd.DataFrame:
    """Lee un archivo Reporte_Por_Principio_Activo_*.xlsx y devuelve un DataFrame
    con columnas estandarizadas más la columna PrincipioActivo."""
    principio_activo = extraer_principio_activo(ruta_archivo)
    frames = []

    try:
        excel = pd.ExcelFile(ruta_archivo, engine="openpyxl")
    except Exception as exc:
        logger.error("No se pudo abrir '%s': %s", ruta_archivo, exc)
        return pd.DataFrame()

    for sheet_name in excel.sheet_names:
        try:
            # Detectar la fila de encabezado buscando "Numero" en la primera columna
            raw = pd.read_excel(
                excel,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
            )
            header_idx = None
            for idx, row in raw.iterrows():
                first_cell = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
                if first_cell == "numero":
                    header_idx = idx
                    break

            if header_idx is None:
                logger.warning(
                    "Hoja '%s' en '%s': no se encontró fila de encabezado, se omite.",
                    sheet_name,
                    os.path.basename(ruta_archivo),
                )
                continue

            df = pd.read_excel(
                excel,
                sheet_name=sheet_name,
                header=header_idx,
                dtype=str,
            )

            # Renombrar columnas al estándar
            rename_map = {
                col: _normalizar_nombre_columna(col) for col in df.columns
            }
            df.rename(columns=rename_map, inplace=True)

            # Conservar solo filas con ComponenteMolecular no nulo
            if "ComponenteMolecular" not in df.columns:
                logger.warning(
                    "Hoja '%s' en '%s': columna 'Componente Molecular' no encontrada.",
                    sheet_name,
                    os.path.basename(ruta_archivo),
                )
                continue

            df = df[df["ComponenteMolecular"].notna()].copy()
            df = df[df["ComponenteMolecular"].str.strip() != ""].copy()

            # Normalizar componente molecular
            df["ComponenteMolecular"] = df["ComponenteMolecular"].apply(normalizar_componente)

            # Agregar columnas de identificación
            df.insert(0, "PrincipioActivo", principio_activo)
            df["FuenteArchivo"] = os.path.basename(ruta_archivo)
            df["HojaOrigen"] = sheet_name

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
# Lectura del análisis de mercado por sector
# ---------------------------------------------------------------------------

def leer_analisis_mercado(ruta_archivo: str) -> pd.DataFrame:
    """Lee AnalisisDelMercadoSector.xlsx y devuelve un DataFrame normalizado."""
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
                    "Hoja '%s' en '%s': no se encontró fila de encabezado.",
                    sheet_name,
                    os.path.basename(ruta_archivo),
                )
                continue

            df = pd.read_excel(excel, sheet_name=sheet_name, header=header_idx, dtype=str)

            rename_map = {col: _normalizar_nombre_columna(col) for col in df.columns}
            df.rename(columns=rename_map, inplace=True)

            # Conservar solo filas con Presentacion no nula ni "TOTALES"
            if "Presentacion" in df.columns:
                df = df[df["Presentacion"].notna()].copy()
                df = df[~df["Presentacion"].str.strip().str.upper().isin(["TOTALES", ""])].copy()
                df["Presentacion"] = df["Presentacion"].str.strip()

            df["FuenteArchivo"] = os.path.basename(ruta_archivo)
            df["HojaOrigen"] = sheet_name

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
# Normalización de nombres de columnas
# ---------------------------------------------------------------------------

_COLUMN_MAP = {
    "numero": "Numero",
    "componente molecular": "ComponenteMolecular",
    "componentemolecular": "ComponenteMolecular",
    "primer trimestre": "PrimerTrimestre",
    "primertrimestre": "PrimerTrimestre",
    "segundo trimestre": "SegundoTrimestre",
    "segundotrimestre": "SegundoTrimestre",
    "tercer trimestre": "TercerTrimestre",
    "tercertrimestre": "TercerTrimestre",
    "cuarto trimestre": "CuartoTrimestre",
    "cuartotrimestre": "CuartoTrimestre",
    "total ventas": "TotalVentas",
    "totalventas": "TotalVentas",
    "presentacion": "Presentacion",
    "empresa": "Empresa",
    "laboratorio": "Laboratorio",
    "rank": "Rank",
    "detalles 2022": "Detalles2022",
    "detalles 2023": "Detalles2023",
    "detalles 2024": "Detalles2024",
    "detalles trimestre i": "DetallesTrimestreI",
    "detalles trimestre ii": "DetallesTrimestreII",
    "detalles trimestre iii": "DetallesTrimestreIII",
    "detalles trimestre iv": "DetallesTrimestreIV",
    "participacion 2022": "Participacion2022",
    "participacion 2023": "Participacion2023",
    "participacion 2024": "Participacion2024",
    "tasa de crecimiento": "TasaCrecimiento",
    "tasa de crecimiento": "TasaCrecimiento",
}


def _normalizar_nombre_columna(nombre: str) -> str:
    """Mapea nombres de columnas a nombres estandarizados."""
    if not isinstance(nombre, str):
        return str(nombre)
    key = normalizar_texto(nombre).lower()
    # Buscar coincidencia exacta
    if key in _COLUMN_MAP:
        return _COLUMN_MAP[key]
    # Buscar coincidencia parcial
    for patron, estandar in _COLUMN_MAP.items():
        if patron in key or key in patron:
            return estandar
    # Sin coincidencia: devolver camelCase simple quitando espacios
    return re.sub(r"\s+", "", nombre.title())


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def consolidar():
    logger.info("=== Consolidador de Principios Activos del Mercado 2026 ===")

    # 1. Leer todos los archivos Reporte_Por_Principio_Activo_2024_*.xlsx
    patron_reporte = os.path.join(SCRIPT_DIR, "Reporte_Por_Principio_Activo_2024_*.xlsx")
    archivos_reporte = sorted(glob.glob(patron_reporte))
    logger.info("Archivos Reporte encontrados: %d", len(archivos_reporte))

    dfs_reporte = []
    for archivo in archivos_reporte:
        logger.info("  Leyendo: %s", os.path.basename(archivo))
        df = leer_reporte(archivo)
        if not df.empty:
            dfs_reporte.append(df)

    df_principios = pd.concat(dfs_reporte, ignore_index=True) if dfs_reporte else pd.DataFrame()
    logger.info("Filas consolidadas de Reportes: %d", len(df_principios))

    # Limpiar: eliminar duplicados exactos
    if not df_principios.empty:
        antes = len(df_principios)
        df_principios.drop_duplicates(
            subset=["PrincipioActivo", "ComponenteMolecular"],
            inplace=True,
        )
        logger.info(
            "Duplicados eliminados (PrincipioActivo+ComponenteMolecular): %d → %d",
            antes,
            len(df_principios),
        )

    # 2. Leer AnalisisDelMercadoSector.xlsx desde carpeta Mercado
    archivos_mercado = sorted(glob.glob(os.path.join(MERCADO_DIR, "*.xlsx")))
    logger.info("Archivos en carpeta Mercado: %d", len(archivos_mercado))

    dfs_mercado = []
    for archivo in archivos_mercado:
        logger.info("  Leyendo: %s", os.path.basename(archivo))
        df = leer_analisis_mercado(archivo)
        if not df.empty:
            dfs_mercado.append(df)

    df_mercado = pd.concat(dfs_mercado, ignore_index=True) if dfs_mercado else pd.DataFrame()
    logger.info("Filas consolidadas del Mercado: %d", len(df_mercado))

    if not df_mercado.empty:
        antes = len(df_mercado)
        df_mercado.drop_duplicates(subset=["Presentacion"], inplace=True)
        logger.info(
            "Duplicados eliminados (Presentacion): %d → %d",
            antes,
            len(df_mercado),
        )

    # 3. Generar el archivo consolidado con múltiples hojas
    if df_principios.empty and df_mercado.empty:
        logger.error("No se encontró ningún dato para consolidar. Saliendo.")
        return

    logger.info("Generando '%s' ...", OUTPUT_FILE)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

        # --- Hoja 1: Principios Activos y Componentes Moleculares ---
        if not df_principios.empty:
            # Columnas prioritarias al frente
            cols_front = [
                "PrincipioActivo",
                "ComponenteMolecular",
                "TotalVentas",
                "PrimerTrimestre",
                "SegundoTrimestre",
                "TercerTrimestre",
                "CuartoTrimestre",
            ]
            cols_extra = [c for c in df_principios.columns if c not in cols_front]
            cols_orden = [c for c in cols_front if c in df_principios.columns] + cols_extra
            df_principios[cols_orden].to_excel(
                writer,
                sheet_name="PrincipiosActivos",
                index=False,
            )
            logger.info(
                "  Hoja 'PrincipiosActivos': %d filas, %d columnas",
                len(df_principios),
                len(cols_orden),
            )

        # --- Hoja 2: Análisis del Mercado por Sector ---
        if not df_mercado.empty:
            cols_front_m = [
                "Presentacion",
                "Empresa",
                "Laboratorio",
                "Detalles2024",
                "Detalles2023",
                "Detalles2022",
                "Participacion2024",
                "TasaCrecimiento",
            ]
            cols_extra_m = [c for c in df_mercado.columns if c not in cols_front_m]
            cols_orden_m = [c for c in cols_front_m if c in df_mercado.columns] + cols_extra_m
            df_mercado[cols_orden_m].to_excel(
                writer,
                sheet_name="MercadoSector",
                index=False,
            )
            logger.info(
                "  Hoja 'MercadoSector': %d filas, %d columnas",
                len(df_mercado),
                len(cols_orden_m),
            )

        # --- Hoja 3: Resumen por PrincipioActivo ---
        if not df_principios.empty and "TotalVentas" in df_principios.columns:
            df_principios["TotalVentas_num"] = pd.to_numeric(
                df_principios["TotalVentas"], errors="coerce"
            )
            resumen = (
                df_principios.groupby("PrincipioActivo", sort=True)
                .agg(
                    CantidadComponentes=("ComponenteMolecular", "count"),
                    TotalVentas=("TotalVentas_num", "sum"),
                )
                .reset_index()
                .sort_values("TotalVentas", ascending=False)
            )
            resumen.to_excel(writer, sheet_name="ResumenPorPrincipioActivo", index=False)
            logger.info(
                "  Hoja 'ResumenPorPrincipioActivo': %d filas", len(resumen)
            )

    logger.info("Archivo generado exitosamente: %s", OUTPUT_FILE)


if __name__ == "__main__":
    consolidar()
