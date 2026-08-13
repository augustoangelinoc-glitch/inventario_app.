# -*- coding: utf-8 -*-
"""
logic.py - Lógica completa con todas las mejoras:
- Parámetros por Familia/Producto
- Preservación de ceros en códigos
- Detección de anomalías
- Pronóstico automático (SES, Holt, Holt-Winters)
- Rotura de stock
- Comentarios explicativos
- Exportación Excel con 11 hojas
"""

import re
import unicodedata
from datetime import datetime
import json

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Utilidades generales
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    """minúsculas, sin tildes, sin espacios/símbolos extra -> para comparar nombres de columnas"""
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return text


# Alias aceptados por cada campo lógico (se comparan ya normalizados)
STOCK_ALIASES = {
    "codigo": ["codigo", "cod", "sku", "item", "codigo producto", "cod producto", "id"],
    "descripcion": ["descripcion", "desc", "nombre", "producto", "nombre producto", "detalle"],
    "familia": ["familia", "categoria", "grupo", "linea", "clase"],
    "unidad": ["unidad de medida", "u m", "um", "unidad", "unid", "medida", "u.m."],
    "stock_actual": ["stock actual", "stock", "cantidad", "existencia", "saldo", "existencias"],
}

SALIDAS_ALIASES = {
    "codigo": STOCK_ALIASES["codigo"],
    "descripcion": STOCK_ALIASES["descripcion"],
    "familia": STOCK_ALIASES["familia"],
    "unidad": STOCK_ALIASES["unidad"],
    "fecha": ["fecha", "fecha salida", "date", "fecha de salida", "fecha movimiento"],
    "cantidad_salida": [
        "cantidad salida", "cantidad de salida", "salida", "cantidad", "qty",
        "cant salida", "unidades", "consumo",
    ],
}


def _find_column(columns, aliases):
    """Busca en `columns` (lista de nombres reales) la que mejor calce con `aliases`."""
    norm_map = {_normalize(c): c for c in columns}
    # 1) match exacto normalizado
    for alias in aliases:
        na = _normalize(alias)
        if na in norm_map:
            return norm_map[na]
    # 2) match parcial (la columna contiene el alias o viceversa)
    for alias in aliases:
        na = _normalize(alias)
        for norm_col, real_col in norm_map.items():
            if na and (na in norm_col or norm_col in na):
                return real_col
    return None


def detect_columns(df: pd.DataFrame, kind: str):
    """kind: 'stock' o 'salidas'. Devuelve dict campo_logico -> nombre_columna_real (o None)."""
    aliases = STOCK_ALIASES if kind == "stock" else SALIDAS_ALIASES
    result = {}
    for field, alist in aliases.items():
        result[field] = _find_column(df.columns, alist)
    return result


def rename_to_standard(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Renombra las columnas detectadas a nombres estándar internos."""
    rename = {v: k for k, v in mapping.items() if v is not None}
    return df.rename(columns=rename)


# ---------------------------------------------------------------------------
# Carga de parámetros por Familia/Producto
# ---------------------------------------------------------------------------

def load_parameters_from_excel(file, tipo="familia") -> dict:
    """
    Carga parámetros desde un archivo Excel.
    tipo: 'familia' o 'producto'
    """
    default_params = {
        "default": {
            "lead_time": 7,
            "dias_cobertura_objetivo": 15,
            "pct_seguridad": 20,
            "metodo_calculo": "Todo el histórico",
            "n_meses": 3
        }
    }
    
    if file is None:
        return default_params
    
    df = pd.read_excel(file)
    params = {}
    
    # Detectar columnas
    col_lead = _find_column(df.columns, ["lead time", "lead_time", "lead"])
    col_cobertura = _find_column(df.columns, ["dias cobertura", "dias_cobertura", "cobertura"])
    col_seguridad = _find_column(df.columns, ["% stock seguridad", "pct_seguridad", "seguridad"])
    col_metodo = _find_column(df.columns, ["metodo calculo", "metodo_calculo", "metodo"])
    col_nmeses = _find_column(df.columns, ["n meses", "n_meses", "meses"])
    
    if tipo == "familia":
        col_key = _find_column(df.columns, ["familia", "categoria", "grupo"])
    else:
        col_key = _find_column(df.columns, ["codigo", "cod", "sku"])
    
    if col_key is None:
        return default_params
    
    for _, row in df.iterrows():
        key = str(row[col_key]).strip()
        if key == "" or key == "nan":
            continue
        params[key] = {
            "lead_time": float(row.get(col_lead, default_params["default"]["lead_time"])),
            "dias_cobertura_objetivo": float(row.get(col_cobertura, default_params["default"]["dias_cobertura_objetivo"])),
            "pct_seguridad": float(row.get(col_seguridad, default_params["default"]["pct_seguridad"])),
            "metodo_calculo": str(row.get(col_metodo, default_params["default"]["metodo_calculo"])),
            "n_meses": int(row.get(col_nmeses, default_params["default"]["n_meses"]))
        }
    
    # Si no hay parámetros personalizados, usar default
    if not params:
        return default_params
    
    # Asegurar que default siempre exista
    if "default" not in params:
        params["default"] = default_params["default"]
    
    return params


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
            "No se pudieron identificar las columnas obligatorias en el archivo de Stock: "
            + ", ".join(missing)
            + ". Verifica que el archivo tenga columnas de Código y Stock Actual."
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
            "No se pudieron identificar las columnas obligatorias en el archivo de Salidas: "
            + ", ".join(missing)
            + ". Verifica que el archivo tenga columnas de Código, Fecha y Cantidad Salida."
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
    """Devuelve un dict con listas de inconsistencias encontradas."""
    issues = {}

    codigos_stock = set(stock_df["codigo"].dropna().unique())
    codigos_salidas = set(salidas_df["codigo"].dropna().unique())

    # Productos en stock sin salidas
    sin_salidas = sorted(codigos_stock - codigos_salidas)
    issues["stock_sin_salidas"] = stock_df[stock_df["codigo"].isin(sin_salidas)][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ].drop_duplicates()

    # Productos con salidas que no están en stock
    sin_stock = sorted(codigos_salidas - codigos_stock)
    issues["salidas_sin_stock"] = salidas_df[salidas_df["codigo"].isin(sin_stock)][
        ["codigo", "descripcion", "familia", "unidad"]
    ].drop_duplicates()

    # Códigos duplicados en stock
    dup_mask = stock_df["codigo"].duplicated(keep=False)
    issues["codigos_duplicados"] = stock_df[dup_mask].sort_values("codigo")

    # Código sin descripción
    issues["sin_descripcion"] = stock_df[stock_df["descripcion"].isna()][
        ["codigo", "familia", "unidad"]
    ].drop_duplicates()

    # Productos sin familia
    issues["sin_familia"] = stock_df[stock_df["familia"].isna()][
        ["codigo", "descripcion", "unidad"]
    ].drop_duplicates()

    # Productos sin unidad de medida
    issues["sin_unidad"] = stock_df[stock_df["unidad"].isna()][
        ["codigo", "descripcion", "familia"]
    ].drop_duplicates()

    # Diferencias de familia entre archivos
    fam_stock = stock_df.dropna(subset=["familia"])[["codigo", "familia"]].drop_duplicates()
    fam_sal = salidas_df.dropna(subset=["familia"])[["codigo", "familia"]].drop_duplicates()
    merged_fam = fam_stock.merge(fam_sal, on="codigo", suffixes=("_stock", "_salidas"))
    issues["diferencia_familia"] = merged_fam[
        merged_fam["familia_stock"] != merged_fam["familia_salidas"]
    ]

    # Diferencias de unidad de medida entre archivos
    um_stock = stock_df.dropna(subset=["unidad"])[["codigo", "unidad"]].drop_duplicates()
    um_sal = salidas_df.dropna(subset=["unidad"])[["codigo", "unidad"]].drop_duplicates()
    merged_um = um_stock.merge(um_sal, on="codigo", suffixes=("_stock", "_salidas"))
    issues["diferencia_unidad"] = merged_um[
        merged_um["unidad_stock"] != merged_um["unidad_salidas"]
    ]

    # Datos vacíos (stock actual nulo)
    issues["stock_vacio"] = stock_df[stock_df["stock_actual"].isna()][
        ["codigo", "descripcion", "familia", "unidad"]
    ]

    # Fechas incorrectas en salidas
    issues["fechas_incorrectas"] = salidas_df[salidas_df["fecha"].isna()][
        ["codigo", "descripcion", "fecha_original"]
    ].rename(columns={"fecha_original": "valor_original"})

    # Cantidades inválidas (nulas o negativas)
    cant_invalidas = salidas_df[
        salidas_df["cantidad_salida"].isna() | (salidas_df["cantidad_salida"] < 0)
    ]
    issues["cantidades_invalidas"] = cant_invalidas[["codigo", "descripcion", "fecha", "cantidad_salida"]]

    # Rotura de stock: stock ≤ 0 pero con salidas
    codigos_rotura = salidas_df[
        salidas_df["codigo"].isin(stock_df[stock_df["stock_actual"] <= 0]["codigo"])
    ]["codigo"].unique()
    issues["rotura_stock"] = stock_df[stock_df["codigo"].isin(codigos_rotura)][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ]

    # Stock negativo
    issues["stock_negativo"] = stock_df[stock_df["stock_actual"] < 0][
        ["codigo", "descripcion", "familia", "unidad", "stock_actual"]
    ]

    total_issues = sum(len(v) for v in issues.values())
    issues["_total"] = total_issues
    return issues


# ---------------------------------------------------------------------------
# Análisis histórico de consumo y pronóstico
# ---------------------------------------------------------------------------

def build_monthly_consumption(salidas_df: pd.DataFrame) -> pd.DataFrame:
    """Consumo por código y mes (periodo calendario)."""
    df = salidas_df.dropna(subset=["fecha", "cantidad_salida"]).copy()
    df = df[df["cantidad_salida"] >= 0]
    df["periodo"] = df["fecha"].dt.to_period("M")
    monthly = (
        df.groupby(["codigo", "periodo"], as_index=False)["cantidad_salida"]
        .sum()
        .rename(columns={"cantidad_salida": "consumo"})
    )
    monthly["periodo_str"] = monthly["periodo"].astype(str)
    return monthly


# ---- Funciones de pronóstico ----

def _forecast_ses(series: np.ndarray, alpha: float = 0.3) -> float:
    """Suavizado Exponencial Simple."""
    if len(series) == 0:
        return 0
    if len(series) == 1:
        return series[0]
    forecast = series[0]
    for value in series[1:]:
        forecast = alpha * value + (1 - alpha) * forecast
    return forecast

def _forecast_holt(series: np.ndarray, alpha: float = 0.3, beta: float = 0.2) -> float:
    """Holt (con tendencia)."""
    if len(series) < 3:
        return series.mean() if len(series) > 0 else 0
    level = series[0]
    trend = series[1] - series[0]
    for i in range(2, len(series)):
        value = series[i]
        level_prev = level
        level = alpha * value + (1 - alpha) * (level_prev + trend)
        trend = beta * (level - level_prev) + (1 - beta) * trend
    return level + trend

def _forecast_hw(series: np.ndarray, alpha: float = 0.3, beta: float = 0.2,
                gamma: float = 0.1, seasonality: int = 12) -> float:
    """Holt-Winters (estacional)."""
    if len(series) < seasonality * 2:
        return _forecast_holt(series, alpha, beta)
    n = len(series)
    seasonals = [series[i] for i in range(seasonality)]
    level = np.mean(series[:seasonality])
    if n >= seasonality * 2:
        trend = (np.mean(series[seasonality:seasonality*2]) - np.mean(series[:seasonality])) / seasonality
    else:
        trend = 0
    for i in range(seasonality, n):
        value = series[i]
        s_idx = i % seasonality
        level_prev = level
        level = alpha * (value / seasonals[s_idx]) + (1 - alpha) * (level_prev + trend)
        trend = beta * (level - level_prev) + (1 - beta) * trend
        seasonals[s_idx] = gamma * (value / level) + (1 - gamma) * seasonals[s_idx]
    next_idx = n % seasonality
    return (level + trend) * seasonals[next_idx]

def _calculate_mape(actual: np.ndarray, forecast: float) -> float:
    """Error Porcentual Absoluto Medio."""
    if len(actual) == 0:
        return 0
    valid = actual > 0
    if not valid.any():
        return 0
    return np.mean(np.abs((actual[valid] - forecast) / actual[valid])) * 100

def _calculate_rmse(actual: np.ndarray, forecast: float) -> float:
    """Raíz del Error Cuadrático Medio."""
    if len(actual) == 0:
        return 0
    return np.sqrt(np.mean((actual - forecast) ** 2))

def _calculate_mae(actual: np.ndarray, forecast: float) -> float:
    """Error Absoluto Medio."""
    if len(actual) == 0:
        return 0
    return np.mean(np.abs(actual - forecast))

def compare_forecast_methods(series: np.ndarray) -> dict:
    """
    Compara todos los métodos de pronóstico y selecciona el mejor.
    """
    if len(series) < 3:
        return {
            "best_method": "promedio",
            "best_params": {},
            "forecast": np.mean(series) if len(series) > 0 else 0,
            "metrics": {"promedio": {"mape": 0, "rmse": 0, "mae": 0}}
        }
    
    n = len(series)
    train_size = int(n * 0.8)
    train = series[:train_size]
    test = series[train_size:]
    
    if len(test) < 1:
        train = series
        test = series[-2:]
    
    methods = {}
    
    # 1. Promedio
    forecast_mean = np.mean(train)
    methods["promedio"] = {
        "forecast": forecast_mean,
        "params": {},
        "mape": _calculate_mape(test, forecast_mean),
        "rmse": _calculate_rmse(test, forecast_mean),
        "mae": _calculate_mae(test, forecast_mean)
    }
    
    # 2. SES
    best_ses = {"mape": float("inf"), "params": {}, "forecast": 0}
    for alpha in [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
        forecast_ses = _forecast_ses(train, alpha)
        mape = _calculate_mape(test, forecast_ses)
        if mape < best_ses["mape"]:
            best_ses = {
                "mape": mape,
                "params": {"alpha": alpha},
                "forecast": forecast_ses,
                "rmse": _calculate_rmse(test, forecast_ses),
                "mae": _calculate_mae(test, forecast_ses)
            }
    methods["ses"] = best_ses
    
    # 3. Holt
    best_holt = {"mape": float("inf"), "params": {}, "forecast": 0}
    for alpha in [0.1, 0.2, 0.3, 0.4, 0.5]:
        for beta in [0.05, 0.1, 0.2, 0.3]:
            try:
                forecast_holt = _forecast_holt(train, alpha, beta)
                mape = _calculate_mape(test, forecast_holt)
                if mape < best_holt["mape"]:
                    best_holt = {
                        "mape": mape,
                        "params": {"alpha": alpha, "beta": beta},
                        "forecast": forecast_holt,
                        "rmse": _calculate_rmse(test, forecast_holt),
                        "mae": _calculate_mae(test, forecast_holt)
                    }
            except:
                continue
    methods["holt"] = best_holt
    
    # 4. Holt-Winters
    if len(series) >= 12:
        best_hw = {"mape": float("inf"), "params": {}, "forecast": 0}
        for alpha in [0.1, 0.2, 0.3, 0.4, 0.5]:
            for beta in [0.05, 0.1, 0.2, 0.3]:
                for gamma in [0.05, 0.1, 0.2]:
                    for seasonality in [12, 6, 3]:
                        try:
                            forecast_hw = _forecast_hw(train, alpha, beta, gamma, seasonality)
                            mape = _calculate_mape(test, forecast_hw)
                            if mape < best_hw["mape"]:
                                best_hw = {
                                    "mape": mape,
                                    "params": {
                                        "alpha": alpha, "beta": beta,
                                        "gamma": gamma, "seasonality": seasonality
                                    },
                                    "forecast": forecast_hw,
                                    "rmse": _calculate_rmse(test, forecast_hw),
                                    "mae": _calculate_mae(test, forecast_hw)
                                }
                        except:
                            continue
        methods["hw"] = best_hw
    
    best_method = min(methods.keys(), key=lambda m: methods[m]["mape"])
    
    return {
        "best_method": best_method,
        "best_params": methods[best_method]["params"],
        "forecast": methods[best_method]["forecast"],
        "metrics": {
            m: {
                "mape": methods[m]["mape"],
                "rmse": methods[m]["rmse"],
                "mae": methods[m]["mae"]
            }
            for m in methods
        }
    }


def compute_consumption_metrics(
    salidas_df: pd.DataFrame,
    metodo: str,
    n_meses: int,
    exclude_anomalies: bool = True,
    anomaly_threshold: float = 2.5,
    forecast_method: str = "auto",
    forecast_params: dict = None
) -> pd.DataFrame:
    """
    Calcula métricas de consumo y pronóstico por producto.
    forecast_method: 'auto' (selecciona el mejor método automáticamente)
    """
    df = salidas_df.dropna(subset=["fecha", "cantidad_salida"]).copy()
    df = df[df["cantidad_salida"] >= 0]

    if df.empty:
        return pd.DataFrame(
            columns=[
                "codigo", "consumo_total", "dias_analizados", "consumo_diario",
                "consumo_semanal", "consumo_mensual", "num_meses",
                "mes_mayor_consumo", "mes_menor_consumo", "tendencia", "variacion_pct",
                "consumo_mensual_depurado", "meses_excluidos",
                "pronostico_proximo_mes", "metodo_pronostico",
                "parametros_pronostico", "mape_pronostico",
                "rmse_pronostico", "mejor_metodo", "comparacion_metodos"
            ]
        )

    monthly = build_monthly_consumption(df)

    rows = []
    for codigo, g in df.groupby("codigo"):
        g = g.sort_values("fecha")
        gm = monthly[monthly["codigo"] == codigo].sort_values("periodo")

        fecha_min, fecha_max = g["fecha"].min(), g["fecha"].max()
        dias_analizados = max((fecha_max - fecha_min).days + 1, 1)
        consumo_total_hist = g["cantidad_salida"].sum()
        num_meses = gm["periodo"].nunique() if not gm.empty else 1
        num_meses = max(num_meses, 1)

        # --- Detectar anomalías para este producto ---
        meses_excluidos = []
        gm_depurado = gm.copy()
        
        if exclude_anomalies and len(gm) >= 3:
            mean_consumo = gm["consumo"].mean()
            std_consumo = gm["consumo"].std()
            
            if std_consumo > 0:
                # Detectar picos estadísticos
                for _, row in gm.iterrows():
                    z_score = abs((row["consumo"] - mean_consumo) / std_consumo)
                    if z_score > anomaly_threshold:
                        meses_excluidos.append(row["periodo_str"])
                        gm_depurado = gm_depurado[gm_depurado["periodo_str"] != row["periodo_str"]]
            
            # Detectar meses con una sola salida que represente > 50% del consumo
            for periodo in gm["periodo_str"].unique():
                salidas_mes = g[g["fecha"].dt.to_period("M").astype(str) == periodo]
                if len(salidas_mes) == 1:
                    consumo_mes = salidas_mes["cantidad_salida"].sum()
                    if consumo_mes > (gm_depurado["consumo"].mean() * 0.5):
                        if periodo not in meses_excluidos:
                            meses_excluidos.append(periodo)
                            gm_depurado = gm_depurado[gm_depurado["periodo_str"] != periodo]

        # --- Calcular consumo ---
        if metodo == "Últimos N meses" and not gm_depurado.empty:
            ult_periodos = gm_depurado["periodo"].sort_values().unique()[-n_meses:]
            gm_sel = gm_depurado[gm_depurado["periodo"].isin(ult_periodos)]
            consumo_mensual = gm_sel["consumo"].mean() if not gm_sel.empty else 0
            consumo_total_periodo = gm_sel["consumo"].sum()
            meses_periodo = max(len(ult_periodos), 1)
            consumo_diario = consumo_total_periodo / (meses_periodo * 30)
        elif metodo == "Promedio ponderado (recientes pesan más)" and not gm_depurado.empty:
            gm_sorted = gm_depurado.sort_values("periodo").reset_index(drop=True)
            pesos = np.arange(1, len(gm_sorted) + 1)
            consumo_mensual = np.average(gm_sorted["consumo"], weights=pesos)
            consumo_diario = consumo_mensual / 30
        else:  # Todo el histórico
            if exclude_anomalies and not gm_depurado.empty:
                consumo_mensual = gm_depurado["consumo"].mean()
                consumo_diario = gm_depurado["consumo"].sum() / (len(gm_depurado) * 30)
            else:
                consumo_mensual = consumo_total_hist / num_meses
                consumo_diario = consumo_total_hist / dias_analizados

        consumo_semanal = consumo_diario * 7
        consumo_mensual_depurado = gm_depurado["consumo"].mean() if not gm_depurado.empty else consumo_mensual

        if not gm.empty:
            mes_mayor = gm.loc[gm["consumo"].idxmax(), "periodo_str"]
            mes_menor = gm.loc[gm["consumo"].idxmin(), "periodo_str"]
        else:
            mes_mayor = mes_menor = None

        # Tendencia sobre datos depurados
        tendencia = "Estable"
        variacion_pct = 0.0
        if len(gm_depurado) >= 2:
            gm_sorted = gm_depu