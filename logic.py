# -*- coding: utf-8 -*-
"""
logic.py - Lógica completa
"""
import re
import unicodedata
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Utilidades generales
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text

STOCK_ALIASES = {
    "codigo": ["codigo", "cod", "sku", "item"],
    "descripcion": ["descripcion", "desc", "nombre"],
    "familia": ["familia", "categoria", "grupo"],
    "unidad": ["unidad de medida", "u m", "um", "unidad"],
    "stock_actual": ["stock actual", "stock", "cantidad", "existencia"],
}

SALIDAS_ALIASES = {
    "codigo": STOCK_ALIASES["codigo"],
    "descripcion": STOCK_ALIASES["descripcion"],
    "familia": STOCK_ALIASES["familia"],
    "unidad": STOCK_ALIASES["unidad"],
    "fecha": ["fecha", "fecha salida", "date"],
    "cantidad_salida": ["cantidad salida", "cantidad", "salida", "qty"],
}

def _find_column(columns, aliases):
    norm_map = {_normalize(c): c for c in columns}
    for alias in aliases:
        na = _normalize(alias)
        if na in norm_map:
            return norm_map[na]
        for norm_col, real_col in norm_map.items():
            if na and (na in norm_col or norm_col in na):
                return real_col
    return None

def detect_columns(df: pd.DataFrame, kind: str):
    aliases = STOCK_ALIASES if kind == "stock" else SALIDAS_ALIASES
    result = {}
    for field, alist in aliases.items():
        result[field] = _find_column(df.columns, alist)
    return result

def rename_to_standard(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    rename = {v: k for k, v in mapping.items() if v is not None}
    return df.rename(columns=rename)

# ---------------------------------------------------------------------------
# Carga de archivos
# ---------------------------------------------------------------------------

def load_stock_file(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype={"Código": str, "codigo": str})
    df.columns = [str(c).strip() for c in df.columns]
    mapping = detect_columns(df, "stock")
    df = rename_to_standard(df, mapping)
    
    missing = [f for f in ["codigo", "stock_actual"] if f not in df.columns]
    if missing:
        raise ValueError(
            "No se pudieron identificar las columnas obligatorias en el archivo de Stock: " + ", ".join(missing)
        )
    
    for opt in ["descripcion", "familia", "unidad"]:
        if opt not in df.columns:
            df[opt] = np.nan

    df["codigo"] = df["codigo"].astype(str).str.strip()
    df["stock_actual"] = pd.to_numeric(df["stock_actual"], errors="coerce")
    df["descripcion"] = df["descripcion"].astype(str).str.strip()
    df["familia"] = df["familia"].astype(str).str.strip()
    df["unidad"] = df["unidad"].astype(str).str.strip().str.upper()
    
    df.loc[df["descripcion"].isin(["nan", "None", ""]), "descripcion"] = np.nan
    df.loc[df["familia"].isin(["nan", "None", ""]), "familia"] = np.nan
    df.loc[df["unidad"].isin(["NAN", "NONE", ""]), "unidad"] = np.nan
    
    return df[["codigo", "descripcion", "familia", "unidad", "stock_actual"]]

def load_salidas_file(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype={"Código": str, "codigo": str})
    df.columns = [str(c).strip() for c in df.columns]
    mapping = detect_columns(df, "salidas")
    df = rename_to_standard(df, mapping)
    
    missing = [f for f in ["codigo", "fecha", "cantidad_salida"] if f not in df.columns]
    if missing:
        raise ValueError(
            "No se pudieron identificar las columnas obligatorias en el archivo de Salidas: " + ", ".join(missing)
        )
    
    for opt in ["descripcion", "familia", "unidad"]:
        if opt not in df.columns:
            df[opt] = np.nan

    df["codigo"] = df["codigo"].astype(str).str.strip()
    df["fecha_original"] = df["fecha"]
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce", dayfirst=True)
    df["cantidad_salida"] = pd.to_numeric(df["cantidad_salida"], errors="coerce")
    df["descripcion"] = df["descripcion"].astype(str).str.strip()
    df["familia"] = df["familia"].astype(str).str.strip()
    df["unidad"] = df["unidad"].astype(str).str.strip().str.upper()
    
    df.loc[df["descripcion"].isin(["nan", "None", ""]), "descripcion"] = np.nan
    df.loc[df["familia"].isin(["nan", "None", ""]), "familia"] = np.nan
    df.loc[df["unidad"].isin(["NAN", "NONE", ""]), "unidad"] = np.nan
    
    return df[["codigo", "descripcion", "familia", "unidad", "fecha", "fecha_original", "cantidad_salida"]]

# ---------------------------------------------------------------------------
# Validación de datos
# ---------------------------------------------------------------------------

def validate_data(stock_df: pd.DataFrame, salidas_df: pd.DataFrame) -> dict:
    issues = {}
    codigos_stock = set(stock_df["codigo"].dropna().unique())
    codigos_salidas = set(salidas_df["codigo"].dropna().unique())

    sin_salidas = sorted(codigos_stock - codigos_salidas)
    issues["stock_sin_salidas"] = stock_df[stock_df["codigo"].isin(sin_salidas)][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ].drop_duplicates()

    sin_stock = sorted(codigos_salidas - codigos_stock)
    issues["salidas_sin_stock"] = salidas_df[salidas_df["codigo"].isin(sin_stock)][
        ["codigo", "descripcion", "familia", "unidad"]
    ].drop_duplicates()

    dup_mask = stock_df["codigo"].duplicated(keep=False)
    issues["codigos_duplicados"] = stock_df[dup_mask].sort_values("codigo")

    issues["sin_descripcion"] = stock_df[stock_df["descripcion"].isna()][
        ["codigo", "familia", "unidad"]
    ].drop_duplicates()

    issues["sin_familia"] = stock_df[stock_df["familia"].isna()][
        ["codigo", "descripcion", "unidad"]
    ].drop_duplicates()

    issues["sin_unidad"] = stock_df[stock_df["unidad"].isna()][
        ["codigo", "descripcion", "familia"]
    ].drop_duplicates()

    fam_stock = stock_df.dropna(subset=["familia"])[["codigo", "familia"]].drop_duplicates()
    fam_sal = salidas_df.dropna(subset=["familia"])[["codigo", "familia"]].drop_duplicates()
    merged_fam = fam_stock.merge(fam_sal, on="codigo", suffixes=("_stock", "_salidas"))
    issues["diferencia_familia"] = merged_fam[
        merged_fam["familia_stock"] != merged_fam["familia_salidas"]
    ]

    um_stock = stock_df.dropna(subset=["unidad"])[["codigo", "unidad"]].drop_duplicates()
    um_sal = salidas_df.dropna(subset=["unidad"])[["codigo", "unidad"]].drop_duplicates()
    merged_um = um_stock.merge(um_sal, on="codigo", suffixes=("_stock", "_salidas"))
    issues["diferencia_unidad"] = merged_um[
        merged_um["unidad_stock"] != merged_um["unidad_salidas"]
    ]

    issues["stock_vacio"] = stock_df[stock_df["stock_actual"].isna()][
        ["codigo", "descripcion", "familia", "unidad"]
    ]

    issues["fechas_incorrectas"] = salidas_df[salidas_df["fecha"].isna()][
        ["codigo", "descripcion", "fecha_original"]
    ].rename(columns={"fecha_original": "valor_original"})

    cant_invalidas = salidas_df[
        salidas_df["cantidad_salida"].isna() | (salidas_df["cantidad_salida"] < 0)
    ]
    issues["cantidades_invalidas"] = cant_invalidas[["codigo", "descripcion", "fecha", "cantidad_salida"]]

    codigos_rotura = salidas_df[
        salidas_df["codigo"].isin(stock_df[stock_df["stock_actual"] <= 0]["codigo"])
    ]["codigo"].unique()
    issues["rotura_stock"] = stock_df[stock_df["codigo"].isin(codigos_rotura)][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ]

    issues["stock_negativo"] = stock_df[stock_df["stock_actual"] < 0][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ]

    total_issues = sum(len(v) for v in issues.values())
    issues["_total"] = total_issues
    return issues

# ---------------------------------------------------------------------------
# Build monthly consumption (Stub)
# ---------------------------------------------------------------------------
def build_monthly_consumption(salidas_df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame()
