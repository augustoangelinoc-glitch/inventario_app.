# -*- coding: utf-8 -*-
"""
logic.py
Lógica de negocio para Control de Inventarios y Abastecimiento.
Compatible con app.py
"""

import io
import re
import unicodedata
import numpy as np
import pandas as pd


# ===========================================================================
# UTILIDADES
# ===========================================================================

def _normalize(text: str) -> str:
    """Convierte texto a minúsculas y elimina tildes."""
    if text is None:
        return ""

    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()

    return text


def _safe_number(value, default=0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def _safe_text(value, default=""):
    if value is None or pd.isna(value):
        return default
    return str(value).strip()


# ===========================================================================
# DETECCIÓN DE COLUMNAS
# ===========================================================================

STOCK_ALIASES = {
    "codigo": [
        "codigo",
        "cod",
        "sku",
        "item",
        "codigo producto",
        "codigo material",
        "material",
    ],
    "descripcion": [
        "descripcion",
        "desc",
        "nombre",
        "descripcion producto",
        "nombre producto",
    ],
    "familia": [
        "familia",
        "categoria",
        "grupo",
        "grupo material",
        "categoria producto",
    ],
    "unidad": [
        "unidad de medida",
        "u m",
        "um",
        "unidad",
        "u.m.",
        "udm",
    ],
    "stock_actual": [
        "stock actual",
        "stock",
        "cantidad",
        "existencia",
        "existencias",
        "inventario",
    ],
}


SALIDAS_ALIASES = {
    "codigo": STOCK_ALIASES["codigo"],
    "descripcion": STOCK_ALIASES["descripcion"],
    "familia": STOCK_ALIASES["familia"],
    "unidad": STOCK_ALIASES["unidad"],
    "fecha": [
        "fecha",
        "fecha salida",
        "fecha de salida",
        "date",
    ],
    "cantidad_salida": [
        "cantidad salida",
        "cantidad",
        "salida",
        "qty",
        "cantidad consumida",
        "consumo",
    ],
}


PARAM_ALIASES = {
    "clave": [
        "familia",
        "producto",
        "codigo",
        "sku",
        "item",
        "material",
        "nombre",
    ],
    "lead_time": [
        "lead time",
        "lead_time",
        "tiempo entrega",
        "dias entrega",
        "plazo entrega",
    ],
    "dias_cobertura_objetivo": [
        "dias cobertura objetivo",
        "dias_cobertura_objetivo",
        "cobertura objetivo",
        "dias cobertura",
    ],
    "pct_seguridad": [
        "pct seguridad",
        "porcentaje seguridad",
        "stock seguridad %",
        "stock seguridad",
        "seguridad",
    ],
    "metodo_calculo": [
        "metodo calculo",
        "metodo_calculo",
        "metodo consumo",
        "metodo",
    ],
    "n_meses": [
        "n meses",
        "n_meses",
        "meses",
        "numero meses",
    ],
}


def _find_column(columns, aliases):
    """Busca una columna por coincidencia exacta o parcial."""
    norm_map = {_normalize(c): c for c in columns}

    for alias in aliases:
        na = _normalize(alias)

        if na in norm_map:
            return norm_map[na]

    for alias in aliases:
        na = _normalize(alias)

        if not na:
            continue

        for norm_col, real_col in norm_map.items():
            if na in norm_col or norm_col in na:
                return real_col

    return None


def detect_columns(df: pd.DataFrame, kind: str):
    if kind == "stock":
        aliases = STOCK_ALIASES
    elif kind == "salidas":
        aliases = SALIDAS_ALIASES
    else:
        aliases = PARAM_ALIASES

    result = {}

    for field, alist in aliases.items():
        result[field] = _find_column(df.columns, alist)

    return result


def rename_to_standard(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    rename = {
        original: standard
        for standard, original in mapping.items()
        if original is not None
    }

    return df.rename(columns=rename)


# ===========================================================================
# CARGA DE STOCK
# ===========================================================================

def load_stock_file(file) -> pd.DataFrame:

    df = pd.read_excel(file)

    df.columns = [str(c).strip() for c in df.columns]

    mapping = detect_columns(df, "stock")

    df = rename_to_standard(df, mapping)

    required = ["codigo", "stock_actual"]

    missing = [x for x in required if x not in df.columns]

    if missing:
        raise ValueError(
            "No se pudieron identificar las columnas obligatorias "
            "del Stock: " + ", ".join(missing)
        )

    for column in ["descripcion", "familia", "unidad"]:
        if column not in df.columns:
            df[column] = np.nan

    df["codigo"] = df["codigo"].astype(str).str.strip()

    df["stock_actual"] = pd.to_numeric(
        df["stock_actual"],
        errors="coerce"
    ).fillna(0)

    for column in ["descripcion", "familia"]:
        df[column] = df[column].astype(str).str.strip()

        df.loc[
            df[column].isin(["nan", "None", ""]),
            column
        ] = np.nan

    df["unidad"] = df["unidad"].astype(str).str.strip().str.upper()

    df.loc[
        df["unidad"].isin(["NAN", "NONE", ""]),
        "unidad"
    ] = np.nan

    return df[
        [
            "codigo",
            "descripcion",
            "familia",
            "unidad",
            "stock_actual",
        ]
    ].copy()


# ===========================================================================
# CARGA DE SALIDAS
# ===========================================================================

def load_salidas_file(file) -> pd.DataFrame:

    df = pd.read_excel(file)

    df.columns = [str(c).strip() for c in df.columns]

    mapping = detect_columns(df, "salidas")

    df = rename_to_standard(df, mapping)

    required = [
        "codigo",
        "fecha",
        "cantidad_salida",
    ]

    missing = [x for x in required if x not in df.columns]

    if missing:
        raise ValueError(
            "No se pudieron identificar las columnas obligatorias "
            "de Salidas: " + ", ".join(missing)
        )

    for column in ["descripcion", "familia", "unidad"]:
        if column not in df.columns:
            df[column] = np.nan

    df["codigo"] = df["codigo"].astype(str).str.strip()

    df["fecha_original"] = df["fecha"]

    df["fecha"] = pd.to_datetime(
        df["fecha"],
        errors="coerce",
        dayfirst=True
    )

    df["cantidad_salida"] = pd.to_numeric(
        df["cantidad_salida"],
        errors="coerce"
    )

    for column in ["descripcion", "familia"]:
        df[column] = df[column].astype(str).str.strip()

        df.loc[
            df[column].isin(["nan", "None", ""]),
            column
        ] = np.nan

    df["unidad"] = df["unidad"].astype(str).str.strip().str.upper()

    df.loc[
        df["unidad"].isin(["NAN", "NONE", ""]),
        "unidad"
    ] = np.nan

    return df[
        [
            "codigo",
            "descripcion",
            "familia",
            "unidad",
            "fecha",
            "fecha_original",
            "cantidad_salida",
        ]
    ].copy()


# ===========================================================================
# PARÁMETROS
# ===========================================================================

def load_parameters_from_excel(file, tipo="familia"):
    """
    Carga parámetros por Familia o por Producto.

    El Excel puede tener columnas como:

    Familia / Código
    Lead Time
    Días Cobertura Objetivo
    % Seguridad
    Método de cálculo
    N meses

    También puede tener una fila DEFAULT.
    """

    df = pd.read_excel(file)

    df.columns = [str(c).strip() for c in df.columns]

    mapping = detect_columns(df, "parametros")

    clave_col = mapping.get("clave")

    if clave_col is None:
        # Intentar identificar automáticamente la primera columna
        # como clave.
        if len(df.columns) > 0:
            clave_col = df.columns[0]
        else:
            raise ValueError(
                "El archivo de parámetros no contiene columnas."
            )

    def get_value(row, field, default):

        col = mapping.get(field)

        if col is None:
            return default

        value = row.get(col, default)

        return value

    resultado = {}

    # Valores por defecto
    default = {
        "lead_time": 7,
        "dias_cobertura_objetivo": 15,
        "pct_seguridad": 20,
        "metodo_calculo": "Todo el histórico",
        "n_meses": 3,
    }

    for _, row in df.iterrows():

        clave = _safe_text(row.get(clave_col, ""))

        if not clave:
            continue

        if _normalize(clave) in [
            "default",
            "defecto",
            "global",
            "general",
        ]:

            default["lead_time"] = _safe_number(
                get_value(row, "lead_time", default["lead_time"]),
                default["lead_time"]
            )

            default["dias_cobertura_objetivo"] = _safe_number(
                get_value(
                    row,
                    "dias_cobertura_objetivo",
                    default["dias_cobertura_objetivo"],
                ),
                default["dias_cobertura_objetivo"]
            )

            default["pct_seguridad"] = _safe_number(
                get_value(
                    row,
                    "pct_seguridad",
                    default["pct_seguridad"],
                ),
                default["pct_seguridad"]
            )

            default["metodo_calculo"] = _safe_text(
                get_value(
                    row,
                    "metodo_calculo",
                    default["metodo_calculo"],
                ),
                default["metodo_calculo"]
            )

            default["n_meses"] = int(
                _safe_number(
                    get_value(
                        row,
                        "n_meses",
                        default["n_meses"],
                    ),
                    default["n_meses"]
                )
            )

            continue

        resultado[clave] = {
            "lead_time": _safe_number(
                get_value(row, "lead_time", default["lead_time"]),
                default["lead_time"]
            ),
            "dias_cobertura_objetivo": _safe_number(
                get_value(
                    row,
                    "dias_cobertura_objetivo",
                    default["dias_cobertura_objetivo"],
                ),
                default["dias_cobertura_objetivo"]
            ),
            "pct_seguridad": _safe_number(
                get_value(
                    row,
                    "pct_seguridad",
                    default["pct_seguridad"],
                ),
                default["pct_seguridad"]
            ),
            "metodo_calculo": _safe_text(
                get_value(
                    row,
                    "metodo_calculo",
                    default["metodo_calculo"],
                ),
                default["metodo_calculo"]
            ),
            "n_meses": int(
                _safe_number(
                    get_value(
                        row,
                        "n_meses",
                        default["n_meses"],
                    ),
                    default["n_meses"]
                )
            ),
        }

    resultado["default"] = default

    return resultado


# ===========================================================================
# VALIDACIÓN
# ===========================================================================

def validate_data(stock_df, salidas_df):

    issues = {}

    stock_codes = set(
        stock_df["codigo"].dropna().astype(str)
    )

    salida_codes = set(
        salidas_df["codigo"].dropna().astype(str)
    )

    # Stock sin salidas
    codes_without_sales = stock_codes - salida_codes

    issues["stock_sin_salidas"] = stock_df[
        stock_df["codigo"].isin(codes_without_sales)
    ].copy()

    # Salidas sin stock
    codes_without_stock = salida_codes - stock_codes

    issues["salidas_sin_stock"] = salidas_df[
        salidas_df["codigo"].isin(codes_without_stock)
    ][
        [
            "codigo",
            "descripcion",
            "familia",
            "unidad",
        ]
    ].drop_duplicates().copy()

    # Duplicados
    duplicated = stock_df[
        stock_df["codigo"].duplicated(
            keep=False
        )
    ].copy()

    issues["codigos_duplicados"] = duplicated

    # Datos faltantes
    issues["sin_descripcion"] = stock_df[
        stock_df["descripcion"].isna()
    ].copy()

    issues["sin_familia"] = stock_df[
        stock_df["familia"].isna()
    ].copy()

    issues["sin_unidad"] = stock_df[
        stock_df["unidad"].isna()
    ].copy()

    # Diferencias
    stock_info = stock_df[
        ["codigo", "familia", "unidad"]
    ].drop_duplicates("codigo")

    salidas_info = salidas_df[
        ["codigo", "familia", "unidad"]
    ].drop_duplicates("codigo")

    merged = stock_info.merge(
        salidas_info,
        on="codigo",
        how="inner",
        suffixes=("_stock", "_salidas")
    )

    issues["diferencia_familia"] = merged[
        merged["familia_stock"].fillna("").astype(str)
        !=
        merged["familia_salidas"].fillna("").astype(str)
    ].copy()

    issues["diferencia_unidad"] = merged[
        merged["unidad_stock"].fillna("").astype(str)
        !=
        merged["unidad_salidas"].fillna("").astype(str)
    ].copy()

    # Stock vacío
    issues["stock_vacio"] = stock_df[
        stock_df["stock_actual"].isna()
    ].copy()

    # Fechas incorrectas
    issues["fechas_incorrectas"] = salidas_df[
        salidas_df["fecha"].isna()
    ].copy()

    # Cantidades inválidas
    issues["cantidades_invalidas"] = salidas_df[
        salidas_df["cantidad_salida"].isna()
        |
        (salidas_df["cantidad_salida"] < 0)
    ].copy()

    # Rotura de stock
    stock_zero = stock_df[
        stock_df["stock_actual"] <= 0
    ]["codigo"]

    issues["rotura_stock"] = salidas_df[
        salidas_df["codigo"].isin(stock_zero)
    ].copy()

    # Stock negativo
    issues["stock_negativo"] = stock_df[
        stock_df["stock_actual"] < 0
    ].copy()

    total = 0

    for key, value in issues.items():
        if isinstance(value, pd.DataFrame):
            total += len(value)

    issues["_total"] = total

    return issues


# ===========================================================================
# CONSUMO MENSUAL
# ===========================================================================

def build_monthly_consumption(salidas_df):

    df = salidas_df.copy()

    df = df[
        df["fecha"].notna()
        &
        df["cantidad_salida"].notna()
    ].copy()

    if df.empty:
        return pd.DataFrame(
            columns=[
                "codigo",
                "descripcion",
                "familia",
                "periodo",
                "periodo_str",
                "consumo",
            ]
        )

    df["periodo"] = df["fecha"].dt.to_period("M")

    monthly = (
        df.groupby(
            [
                "codigo",
                "descripcion",
                "familia",
                "periodo",
            ],
            dropna=False,
            as_index=False
        )["cantidad_salida"]
        .sum()
        .rename(
            columns={
                "cantidad_salida": "consumo"
            }
        )
    )

    monthly["periodo_str"] = (
        monthly["periodo"]
        .astype(str)
    )

    return monthly


# ===========================================================================
# ANOMALÍAS
# ===========================================================================

def _detect_anomalies(series, threshold=2.5):

    values = pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)

    if len(values) < 3:
        return pd.Series(False, index=series.index)

    std = values.std()

    if std == 0 or pd.isna(std):
        return pd.Series(False, index=series.index)

    mean = values.mean()

    return abs(values - mean) > threshold * std


# ===========================================================================
# CONSUMO Y PRONÓSTICO
# ===========================================================================

def compute_consumption_metrics(
    salidas_df,
    metodo_calculo="Todo el histórico",
    n_meses=3,
    exclude_anomalies=True,
    anomaly_threshold=2.5,
    forecast_method="auto",
    forecast_params=None,
):

    monthly = build_monthly_consumption(
        salidas_df
    )

    if monthly.empty:
        return pd.DataFrame(
            columns=[
                "codigo",
                "descripcion",
                "familia",
                "consumo_total",
                "consumo_mensual_historico",
                "consumo_mensual",
            ]
        )

    resultados = []

    for codigo, grupo in monthly.groupby(
        "codigo"
    ):

        grupo = grupo.sort_values("periodo").copy()

        descripcion = (
            grupo["descripcion"]
            .dropna()
            .iloc[0]
            if grupo["descripcion"].notna().any()
            else ""
        )

        familia = (
            grupo["familia"]
            .dropna()
            .iloc[0]
            if grupo["familia"].notna().any()
            else "Sin familia"
        )

        serie_original = grupo[
            "consumo"
        ].astype(float)

        consumo_total = serie_original.sum()

        historico_promedio = (
            serie_original.mean()
            if len(serie_original) > 0
            else 0
        )

        serie_calculo = serie_original.copy()

        meses_excluidos = []

        if exclude_anomalies and len(
            serie_calculo
        ) >= 3:

            flags = _detect_anomalies(
                serie_calculo,
                anomaly_threshold
            )

            if flags.any():

                meses_excluidos = (
                    grupo.loc[
                        flags,
                        "periodo_str"
                    ]
                    .tolist()
                )

                depurada = serie_calculo[
                    ~flags
                ]

                if len(depurada) > 0:
                    serie_calculo = depurada

        # Últimos N meses
        if metodo_calculo == "Últimos N meses":

            serie_calculo = serie_calculo.tail(
                int(n_meses)
            )

        consumo_mensual = (
            serie_calculo.mean()
            if len(serie_calculo) > 0
            else 0
        )

        # Pronóstico sencillo
        metodo_usado = "Promedio"

        if forecast_method == "auto":

            if len(serie_calculo) >= 3:

                recientes = serie_calculo.tail(3)

                pesos = np.arange(
                    1,
                    len(recientes) + 1
                )

                pronostico = np.average(
                    recientes,
                    weights=pesos
                )

                consumo_mensual = (
                    pronostico
                    if np.isfinite(pronostico)
                    else consumo_mensual
                )

                metodo_usado = (
                    "Promedio ponderado"
                )

            else:

                metodo_usado = "Promedio histórico"

        # MAPE aproximado
        if len(serie_calculo) > 1:

            media = serie_calculo.mean()

            if media != 0:
                mape = (
                    np.mean(
                        abs(
                            serie_calculo - media
                        ) / abs(media)
                    )
                    * 100
                )
            else:
                mape = 0
        else:
            mape = 0

        comentario_anomalias = (
            "Sin anomalías detectadas"
            if not meses_excluidos
            else
            "Meses excluidos: "
            + ", ".join(meses_excluidos)
        )

        resultados.append(
            {
                "codigo": str(codigo),
                "descripcion": descripcion,
                "familia": familia,
                "consumo_total": consumo_total,
                "consumo_mensual_historico": historico_promedio,
                "consumo_mensual": consumo_mensual,
                "meses_excluidos": (
                    "Ninguno"
                    if not meses_excluidos
                    else ", ".join(meses_excluidos)
                ),
                "metodo_pronostico_usado": metodo_usado,
                "mape_pronostico": mape,
                "comentario_consumo": (
                    "Consumo calculado con "
                    + metodo_calculo
                ),
                "comentario_pronostico": (
                    f"Método utilizado: {metodo_usado}"
                ),
                "comentario_anomalias":
                    comentario_anomalias,
            }
        )

    return pd.DataFrame(resultados)


# ===========================================================================
# PARÁMETROS POR PRODUCTO/FAMILIA
# ===========================================================================

def _get_parameters(
    params_dict,
    codigo,
    familia,
):

    default = params_dict.get(
        "default",
        {
            "lead_time": 7,
            "dias_cobertura_objetivo": 15,
            "pct_seguridad": 20,
            "metodo_calculo": "Todo el histórico",
            "n_meses": 3,
        }
    )

    # Primero producto
    if codigo is not None:

        codigo = str(codigo)

        if codigo in params_dict:
            return params_dict[codigo]

    # Luego familia
    if familia is not None:

        familia = str(familia)

        if familia in params_dict:
            return params_dict[familia]

    return default


# ===========================================================================
# INDICADORES DE ABASTECIMIENTO
# ===========================================================================

def compute_supply_indicators(
    stock_df,
    consumption_df,
    lead_time=None,
    dias_cobertura_objetivo=None,
    pct_seguridad=None,
    params_dict=None,
    usar_consumo_depurado=True,
    usar_pronostico=True,
):

    stock = stock_df.copy()

    consumo = consumption_df.copy()

    resultado = stock.merge(
        consumo,
        on=[
            "codigo",
            "descripcion",
            "familia",
        ],
        how="outer",
        suffixes=("", "_consumo")
    )

    # Recuperar descripción/familia cuando vengan de consumo
    for col in ["descripcion", "familia"]:

        alt = f"{col}_consumo"

        if alt in resultado.columns:

            resultado[col] = (
                resultado[col]
                .fillna(resultado[alt])
            )

            resultado.drop(
                columns=[alt],
                inplace=True
            )

    resultado["stock_actual"] = pd.to_numeric(
        resultado["stock_actual"],
        errors="coerce"
    ).fillna(0)

    resultado["consumo_mensual"] = pd.to_numeric(
        resultado["consumo_mensual"],
        errors="coerce"
    ).fillna(0)

    resultado["consumo_mensual_historico"] = pd.to_numeric(
        resultado["consumo_mensual_historico"],
        errors="coerce"
    ).fillna(0)

    leads = []
    coberturas = []
    seguridades = []

    for _, row in resultado.iterrows():

        if params_dict is not None:

            p = _get_parameters(
                params_dict,
                row["codigo"],
                row["familia"]
            )

            lt = _safe_number(
                p.get("lead_time", 7),
                7
            )

            dc = _safe_number(
                p.get(
                    "dias_cobertura_objetivo",
                    15
                ),
                15
            )

            ps = _safe_number(
                p.get(
                    "pct_seguridad",
                    20
                ),
                20
            )

        else:

            lt = _safe_number(
                lead_time,
                7
            )

            dc = _safe_number(
                dias_cobertura_objetivo,
                15
            )

            ps = _safe_number(
                pct_seguridad,
                20
            )

        leads.append(lt)
        coberturas.append(dc)
        seguridades.append(ps)

    resultado["lead_time"] = leads

    resultado["dias_cobertura_objetivo"] = coberturas

    resultado["pct_seguridad"] = seguridades

    # Consumo diario
    resultado["consumo_diario"] = (
        resultado["consumo_mensual"] / 30
    )

    # Cobertura
    resultado["dias_cobertura"] = np.where(
        resultado["consumo_diario"] > 0,
        resultado["stock_actual"]
        /
        resultado["consumo_diario"],
        np.inf
    )

    resultado["dias_cobertura_display"] = (
        resultado["dias_cobertura"]
        .replace(np.inf, np.nan)
    )

    # Stock seguridad
    resultado["stock_seguridad"] = (
        resultado["consumo_mensual"]
        *
        resultado["pct_seguridad"]
        / 100
    )

    resultado["punto_pedido"] = (
        resultado["consumo_diario"]
        *
        resultado["lead_time"]
        +
        resultado["stock_seguridad"]
    )

    resultado["stock_objetivo"] = (
        resultado["consumo_diario"]
        *
        resultado["dias_cobertura_objetivo"]
        +
        resultado["stock_seguridad"]
    )

    resultado["cantidad_comprar"] = (
        resultado["stock_objetivo"]
        -
        resultado["stock_actual"]
    ).clip(lower=0)

    # Estados
    def estado(row):

        stock = row["stock_actual"]
        punto = row["punto_pedido"]
        objetivo = row["stock_objetivo"]

        if stock <= 0 and row["consumo_mensual"] > 0:
            return "🔴 CRÍTICO"

        if stock < punto:
            return "🟡 POR ABASTECER"

        if stock > objetivo * 1.5:
            return "🔵 EXCESO"

        return "🟢 OK"

    resultado["estado"] = resultado.apply(
        estado,
        axis=1
    )

    # Prioridad
    def prioridad(row):

        estado = row["estado"]

        if estado == "🔴 CRÍTICO":
            return 1

        if estado == "🟡 POR ABASTECER":
            return 2

        if estado == "🔵 EXCESO":
            return 4

        return 3

    resultado["prioridad_compra"] = resultado.apply(
        prioridad,
        axis=1
    )

    # Comentarios
    resultado["comentario_dias_cobertura"] = (
        "Cobertura calculada según stock actual y consumo diario."
    )

    resultado["comentario_stock_seguridad"] = (
        "Stock de seguridad calculado como porcentaje del consumo mensual."
    )

    resultado["comentario_punto_pedido"] = (
        "Punto de pedido = consumo durante Lead Time + stock de seguridad."
    )

    resultado["comentario_stock_objetivo"] = (
        "Stock objetivo = consumo correspondiente a los días de cobertura + seguridad."
    )

    resultado["recomendacion"] = resultado.apply(
        lambda r:
        (
            f"Se recomienda comprar "
            f"{r['cantidad_comprar']:.0f} "
            f"{r['unidad'] or 'UND'}."
            if r["cantidad_comprar"] > 0
            else
            "No se requiere compra inmediata."
        ),
        axis=1
    )

    resultado["stock_origen"] = resultado[
        "stock_actual"
    ]

    return resultado


# ===========================================================================
# EXPORTACIÓN EXCEL
# ===========================================================================

def build_export_excel(
    resultado_df,
    monthly_df,
    issues,
    params,
    anomalies_df,
    products_added_df,
    output_path,
):

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        resultado_df.to_excel(
            writer,
            sheet_name="Abastecimiento",
            index=False
        )

        monthly_df.to_excel(
            writer,
            sheet_name="Consumo Mensual",
            index=False
        )

        # Validaciones
        for key, value in issues.items():

            if key.startswith("_"):
                continue

            if isinstance(value, pd.DataFrame):

                sheet = key[:31]

                value.to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

        if anomalies_df is not None:

            anomalies_df.to_excel(
                writer,
                sheet_name="Anomalias",
                index=False
            )

        if products_added_df is not None:

            products_added_df.to_excel(
                writer,
                sheet_name="Productos Agregados",
                index=False
            )

        # Parámetros globales
        pd.DataFrame(
            [
                {
                    "Parametro": key,
                    "Valor": value,
                }
                for key, value in params.items()
                if not isinstance(value, (dict, list))
            ]
        ).to_excel(
            writer,
            sheet_name="Parametros",
            index=False
        )

    return output_path
