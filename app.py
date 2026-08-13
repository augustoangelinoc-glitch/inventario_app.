# -*- coding: utf-8 -*-
"""
App de Control de Inventarios y Abastecimiento - Versión Mejorada
Ejecutar con: streamlit run app.py
"""

import io
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from logic import (
    load_stock_file,
    load_salidas_file,
    load_parameters_from_excel,
    validate_data,
    compute_consumption_metrics,
    build_monthly_consumption,
    compute_supply_indicators,
    build_export_excel,
)

st.set_page_config(
    page_title="Control de Inventarios y Abastecimiento",
    page_icon="📦",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Estado de sesión
# ---------------------------------------------------------------------------
if "resultado" not in st.session_state:
    st.session_state.resultado = None
    st.session_state.monthly = None
    st.session_state.issues = None
    st.session_state.stock_df = None
    st.session_state.salidas_df = None
    st.session_state.params_dict = None
    st.session_state.anomalies_df = None
    st.session_state.products_added_df = None

# ---------------------------------------------------------------------------
# Sidebar - Parámetros configurables
# ---------------------------------------------------------------------------
st.sidebar.title("⚙️ Parámetros de Abastecimiento")

# Tipo de configuración
tipo_config = st.sidebar.radio(
    "Configuración de parámetros",
    ["Global", "Por Familia", "Por Producto"],
    help="Global: todos los productos usan los mismos parámetros. Por Familia/Producto: carga un archivo Excel con parámetros personalizados."
)

if tipo_config == "Global":
    lead_time = st.sidebar.number_input("Lead Time (días)", min_value=0, value=7, step=1)
    dias_cobertura_objetivo = st.sidebar.number_input(
        "Días de cobertura objetivo", min_value=0, value=15, step=1
    )
    pct_seguridad = st.sidebar.slider(
        "Porcentaje de stock de seguridad (%)", min_value=0, max_value=200, value=20, step=5
    )
    metodo_calculo = st.sidebar.selectbox(
        "Método de cálculo del consumo",
        ["Todo el histórico", "Últimos N meses", "Promedio ponderado (recientes pesan más)"],
    )
    n_meses = 3
    if metodo_calculo == "Últimos N meses":
        n_meses = st.sidebar.number_input("N° de meses a considerar", min_value=1, value=3, step=1)
    
    params_dict = None
    st.session_state.params_dict = None

elif tipo_config == "Por Familia":
    st.sidebar.info("Carga un archivo Excel con parámetros por Familia")
    params_file = st.sidebar.file_uploader("Archivo de parámetros por Familia", type=["xlsx", "xls"])
    if params_file:
        params_dict = load_parameters_from_excel(params_file, tipo="familia")
        st.session_state.params_dict = params_dict
        st.sidebar.success(f"✅ Parámetros cargados para {len(params_dict)-1} familias")
        lead_time = params_dict["default"]["lead_time"]
        dias_cobertura_objetivo = params_dict["default"]["dias_cobertura_objetivo"]
        pct_seguridad = params_dict["default"]["pct_seguridad"]
        metodo_calculo = params_dict["default"]["metodo_calculo"]
        n_meses = params_dict["default"]["n_meses"]
    else:
        st.sidebar.warning("Sube un archivo de parámetros para continuar")
        params_dict = None

elif tipo_config == "Por Producto":
    st.sidebar.info("Carga un archivo Excel con parámetros por Producto")
    params_file = st.sidebar.file_uploader("Archivo de parámetros por Producto", type=["xlsx", "xls"])
    if params_file:
        params_dict = load_parameters_from_excel(params_file, tipo="producto")
        st.session_state.params_dict = params_dict
        st.sidebar.success(f"✅ Parámetros cargados para {len(params_dict)-1} productos")
        lead_time = params_dict["default"]["lead_time"]
        dias_cobertura_objetivo = params_dict["default"]["dias_cobertura_objetivo"]
        pct_seguridad = params_dict["default"]["pct_seguridad"]
        metodo_calculo = params_dict["default"]["metodo_calculo"]
        n_meses = params_dict["default"]["n_meses"]
    else:
        st.sidebar.warning("Sube un archivo de parámetros para continuar")
        params_dict = None

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Configuración de anomalías")

excluir_anomalias = st.sidebar.checkbox(
    "Excluir meses anómalos del cálculo de compra", 
    value=True,
    help="Si está activado, los meses con consumo atípico (picos o salidas únicas) se excluyen del cálculo de consumo para compras."
)

umbral_anomalia = st.sidebar.slider(
    "Umbral para detectar picos (desviaciones estándar)",
    min_value=1.0, max_value=5.0, value=2.5, step=0.5,
    help="Cuanto más bajo, más sensible para detectar picos de consumo."
)

st.sidebar.markdown("---")
st.sidebar.subheader("📈 Pronóstico automático")

usar_pronostico = st.sidebar.checkbox(
    "Usar pronóstico automático para compras",
    value=True,
    help="El sistema prueba SES, Holt y Holt-Winters y selecciona el mejor método para cada producto."
)

st.sidebar.caption(
    "Estos parámetros se aplican a todos los productos. "
    "Para configuración por Familia o Producto, usa la opción correspondiente."
)

params = {
    "lead_time": lead_time if params_dict is None else None,
    "dias_cobertura_objetivo": dias_cobertura_objetivo if params_dict is None else None,
    "pct_seguridad": pct_seguridad if params_dict is None else None,
    "metodo_calculo": metodo_calculo if params_dict is None else None,
    "n_meses": n_meses if params_dict is None else None,
    "excluir_anomalias": excluir_anomalias,
    "umbral_anomalia": umbral_anomalia,
    "usar_pronostico": usar_pronostico,
}

# ---------------------------------------------------------------------------
# Encabezado
# ---------------------------------------------------------------------------
st.title("📦 Control de Inventarios y Abastecimiento")
st.caption("Sube el Stock Actual y las Salidas Históricas. El sistema calcula automáticamente todo el análisis con pronóstico inteligente.")

# ---------------------------------------------------------------------------
# 1-2. Carga de archivos
# ---------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ Subir Stock Actual")
    stock_file = st.file_uploader("Archivo Excel de Stock Actual", type=["xlsx", "xls"], key="stock_upl")
with col2:
    st.subheader("2️⃣ Subir Salidas Históricas")
    salidas_file = st.file_uploader("Archivo Excel de Salidas Históricas", type=["xlsx", "xls"], key="salidas_upl")

st.subheader("3️⃣ Configurar Parámetros")
st.caption("Los parámetros se configuran en el panel lateral izquierdo (⚙️).")

calc_col, _ = st.columns([1, 3])
with calc_col:
    calcular = st.button("🚀 4️⃣ CALCULAR ABASTECIMIENTO", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# 4. Procesamiento
# ---------------------------------------------------------------------------
if calcular:
    if stock_file is None or salidas_file is None:
        st.error("Debes subir ambos archivos: Stock Actual y Salidas Históricas.")
    else:
        try:
            with st.spinner("Leyendo y validando archivos..."):
                stock_df = load_stock_file(stock_file)
                salidas_df = load_salidas_file(salidas_file)
                issues = validate_data(stock_df, salidas_df)
                
                # Manejar productos con salidas sin stock
                products_added_df = pd.DataFrame()
                if not issues.get("salidas_sin_stock", pd.DataFrame()).empty and st.session_state.get("decision_sin_stock", "") == "":
                    st.warning(f"⚠️ Se detectaron {len(issues['salidas_sin_stock'])} productos con salidas pero sin stock.")
                    decision = st.radio(
                        "¿Qué deseas hacer con estos productos?",
                        [
                            "✅ Crearlos en Stock con stock 0 (recomendado)",
                            "❌ Ignorarlos y no mostrarlos en resultados",
                            "⚠️ Mostrarlos en análisis pero sin recomendación de compra"
                        ],
                        index=0,
                        key="decision_sin_stock_radio"
                    )
                    st.session_state["decision_sin_stock"] = decision
                    
                    if decision.startswith("✅"):
                        nuevos_productos = []
                        for _, row in issues["salidas_sin_stock"].iterrows():
                            nuevos_productos.append({
                                "codigo": row["codigo"],
                                "descripcion": row["descripcion"] if pd.notna(row["descripcion"]) else "Producto sin descripción",
                                "familia": row["familia"] if pd.notna(row["familia"]) else "Sin familia",
                                "unidad": row["unidad"] if pd.notna(row["unidad"]) else "UND",
                                "stock_actual": 0
                            })
                        if nuevos_productos:
                            nuevos_df = pd.DataFrame(nuevos_productos)
                            stock_df = pd.concat([stock_df, nuevos_df], ignore_index=True)
                            products_added_df = nuevos_df
                            st.success(f"✅ Se agregaron {len(nuevos_productos)} productos al Stock con stock 0.")
                            issues["salidas_sin_stock"] = pd.DataFrame()
                            issues["_total"] = sum(len(v) for k, v in issues.items() if not k.startswith("_"))
                    
                    elif decision.startswith("❌"):
                        codigos_excluir = issues["salidas_sin_stock"]["codigo"].tolist()
                        salidas_df = salidas_df[~salidas_df["codigo"].isin(codigos_excluir)]
                        st.info(f"ℹ️ Se excluyeron {len(codigos_excluir)} productos del análisis.")
                        issues["salidas_sin_stock"] = pd.DataFrame()
                        issues["_total"] = sum(len(v) for k, v in issues.items() if not k.startswith("_"))
                
                st.session_state.issues = issues
                st.session_state.stock_df = stock_df
                st.session_state.salidas_df = salidas_df
                st.session_state.products_added_df = products_added_df

            with st.spinner("Calculando consumo histórico, pronóstico e indicadores de abastecimiento..."):
                if st.session_state.params_dict is not None:
                    consumption_df = compute_consumption_metrics(
                        salidas_df, 
                        st.session_state.params_dict["default"]["metodo_calculo"],
                        st.session_state.params_dict["default"]["n_meses"],
                        exclude_anomalies=excluir_anomalias,
                        anomaly_threshold=umbral_anomalia,
                        forecast_method="auto" if usar_pronostico else "promedio",
                        forecast_params={}
                    )
                    resultado_df = compute_supply_indicators(
                        stock_df, 
                        consumption_df, 
                        st.session_state.params_dict,
                        usar_consumo_depurado=excluir_anomalias,
                        usar_pronostico=usar_pronostico
                    )
                else:
                    consumption_df = compute_consumption_metrics(
                        salidas_df, 
                        metodo_calculo, 
                        n_meses,
                        exclude_anomalies=excluir_anomalias,
                        anomaly_threshold=umbral_anomalia,
                        forecast_method="auto" if usar_pronostico else "promedio",
                        forecast_params={}
                    )
                    resultado_df = compute_supply_indicators(
                        stock_df, 
                        consumption_df, 
                        lead_time, 
                        dias_cobertura_objetivo, 
                        pct_seguridad,
                        usar_consumo_depurado=excluir_anomalias,
                        usar_pronostico=usar_pronostico
                    )
                
                monthly_df = build_monthly_consumption(salidas_df)
                
                if "meses_excluidos" in resultado_df.columns:
                    anomalies_df = resultado_df[resultado_df["meses_excluidos"] != "Ninguno"][
                        ["codigo", "descripcion", "familia", "meses_excluidos", "consumo_mensual_historico", "consumo_mensual"]
                    ].copy()
                    anomalies_df.rename(columns={
                        "meses_excluidos": "Meses excluidos",
                        "consumo_mensual_historico": "Consumo original",
                        "consumo_mensual": "Consumo depurado"
                    }, inplace=True)
                else:
                    anomalies_df = pd.DataFrame()

            st.session_state.resultado = resultado_df
            st.session_state.monthly = monthly_df
            st.session_state.anomalies_df = anomalies_df
            st.success("✅ Cálculo completado correctamente.")
            
        except Exception as e:
            st.error(f"❌ Ocurrió un error procesando los archivos: {e}")

resultado_df = st.session_state.resultado

# ---------------------------------------------------------------------------
# Resultados
# ---------------------------------------------------------------------------
if resultado_df is not None:
    issues = st.session_state.issues
    monthly_df = st.session_state.monthly
    stock_df = st.session_state.stock_df
    salidas_df = st.session_state.salidas_df
    anomalies_df = st.session_state.anomalies_df
    products_added_df = st.session_state.products_added_df

    # --- 17. Resumen de carga ---
    st.markdown("---")
    st.subheader("📋 Resumen de Carga")
    meses_disponibles = sorted(monthly_df["periodo_str"].unique()) if not monthly_df.empty else []
    periodo_txt = f"{meses_disponibles[0]} – {meses_disponibles[-1]}" if meses_disponibles else "Sin datos"

    r1, r2, r3, r4, r5, r6 = st.columns(6)
    r1.metric("Stock cargado", f"{len(stock_df):,} productos")
    r2.metric("Salidas cargadas", f"{len(salidas_df):,} registros")
    r3.metric("Familias", f"{resultado_df['familia'].nunique():,}")
    r4.metric("Periodo analizado", periodo_txt)
    r5.metric("Inconsistencias", f"{issues['_total']:,}")
    r6.metric("Pronóstico automático", "✅ Activo" if params.get("usar_pronostico") else "❌ Inactivo")

    # --- Validación de Datos ---
    with st.expander("🔍 Validación de Datos", expanded=(issues["_total"] > 0)):
        etiquetas = {
            "stock_sin_salidas": "Productos en Stock sin salidas registradas",
            "salidas_sin_stock": "Productos con salidas que no aparecen en Stock",
            "codigos_duplicados": "Códigos duplicados en Stock",
            "sin_descripcion": "Códigos sin descripción",
            "sin_familia": "Productos sin Familia",
            "sin_unidad": "Productos sin Unidad de Medida",
            "diferencia_familia": "Diferencias de Familia entre archivos",
            "diferencia_unidad": "Diferencias de Unidad de Medida entre archivos",
            "stock_vacio": "Registros de Stock con datos vacíos",
            "fechas_incorrectas": "Fechas incorrectas en Salidas",
            "cantidades_invalidas": "Cantidades inválidas en Salidas",
            "rotura_stock": "Productos con rotura de stock (stock ≤ 0 con salidas)",
            "stock_negativo": "Stock actual negativo",
        }
        if issues["_total"] == 0:
            st.success("No se encontraron inconsistencias en los datos. ✅")
        else:
            for key, label in etiquetas.items():
                df_issue = issues.get(key)
                if df_issue is not None and not df_issue.empty:
                    st.markdown(f"**{label}** — {len(df_issue)} registro(s)")
                    st.dataframe(df_issue, use_container_width=True, hide_index=True)

    # --- 10. Dashboard ---
    st.markdown("---")
    st.subheader("📊 Dashboard General")

    total_productos = len(resultado_df)
    criticos = (resultado_df["estado"] == "🔴 CRÍTICO").sum()
    por_abastecer = (resultado_df["estado"] == "🟡 POR ABASTECER").sum()
    ok_count = (resultado_df["estado"] == "🟢 OK").sum()
    exceso = (resultado_df["estado"] == "🔵 EXCESO").sum()
    consumo_total = resultado_df["consumo_total"].sum()
    consumo_prom_mensual = resultado_df["consumo_mensual"].sum()
    compra_total = resultado_df["cantidad_comprar"].sum()
    
    if "mape_pronostico" in resultado_df.columns:
        mape_promedio = resultado_df[resultado_df["mape_pronostico"] > 0]["mape_pronostico"].mean()
        if not np.isnan(mape_promedio):
            mape_promedio = f"{mape_promedio:.1f}%"
        else:
            mape_promedio = "N/A"
    else:
        mape_promedio = "N/A"

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total productos", f"{total_productos:,}")
    k2.metric("Familias", f"{resultado_df['familia'].nunique():,}")
    k3.metric("Unidades de medida", f"{resultado_df['unidad'].nunique():,}")
    k4.metric("Consumo total histórico", f"{consumo_total:,.0f}")
    k5.metric("Compra sugerida total", f"{compra_total:,.0f}")
    k6.metric("Error MAPE promedio", mape_promedio)

    e1, e2, e3, e4 = st.columns(4)
    e1.metric("🔴 Críticos", f"{criticos:,}")
    e2.metric("🟡 Por abastecer", f"{por_abastecer:,}")
    e3.metric("🟢 OK", f"{ok_count:,}")
    e4.metric("🔵 Exceso", f"{exceso:,}")

    g1, g2 = st.columns(2)
    with g1:
        if not monthly_df.empty:
            consumo_mes = monthly_df.groupby("periodo_str", as_index=False)["consumo"].sum().sort_values("periodo_str")
            fig = px.bar(consumo_mes, x="periodo_str", y="consumo", title="Consumo Mensual Total",
                         labels={"periodo_str": "Mes", "consumo": "Consumo"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin datos de consumo mensual para graficar.")
    with g2:
        fam_consumo = resultado_df.groupby("familia", as_index=False)["consumo_mensual"].sum()
        fig = px.bar(fam_consumo.sort_values("consumo_mensual", ascending=False), x="familia", y="consumo_mensual",
                     title="Consumo Mensual por Familia", labels={"familia": "Familia", "consumo_mensual": "Consumo Mensual"})
        st.plotly_chart(fig, use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        top20 = resultado_df.nlargest(20, "stock_actual")
        fig = px.bar(
            top20, x="codigo", y=["stock_actual", "stock_objetivo"], barmode="group",
            title="Stock Actual vs. Stock Objetivo (Top 20 por stock)",
            labels={"value": "Cantidad", "codigo": "Código", "variable": "Indicador"},
        )
        st.plotly_chart(fig, use_container_width=True)
    with g4:
        estado_dist = resultado_df["estado"].value_counts().reset_index()
        estado_dist.columns = ["estado", "cantidad"]
        fig = px.pie(estado_dist, names="estado", values="cantidad", title="Distribución por Estado de Abastecimiento")
        st.plotly_chart(fig, use_container_width=True)

    g5, g6 = st.columns(2)
    with g5:
        top_consumo = resultado_df.nlargest(10, "consumo_mensual")
        fig = px.bar(top_consumo.sort_values("consumo_mensual"), x="consumo_mensual", y="codigo", orientation="h",
                     title="Ranking: Productos con Mayor Consumo", labels={"consumo_mensual": "Consumo Mensual", "codigo": "Código"})
        st.plotly_chart(fig, use_container_width=True)
    with g6:
        top_compra = resultado_df.nlargest(10, "cantidad_comprar")
        fig = px.bar(top_compra.sort_values("cantidad_comprar"), x="cantidad_comprar", y="codigo", orientation="h",
                     title="Ranking: Productos que Requieren Mayor Compra", labels={"cantidad_comprar": "Cantidad a Comprar", "codigo": "Código"})
        st.plotly_chart(fig, use_container_width=True)

    criticos_df = resultado_df[resultado_df["estado"] == "🔴 CRÍTICO"]
    if not criticos_df.empty:
        fig = px.bar(
            criticos_df.nlargest(15, "cantidad_comprar").sort_values("cantidad_comprar"),
            x="cantidad_comprar", y="codigo", orientation="h", title="Productos Críticos (Top 15 por cantidad a comprar)",
            labels={"cantidad_comprar": "Cantidad a Comprar", "codigo": "Código"},
        )
        st.plotly_chart(fig, use_container_width=True)

    # --- 8. Anomalías de Consumo ---
    if anomalies_df is not None and not anomalies_df.empty:
        st.markdown("---")
        st.subheader("🔍 Anomalías de Consumo Detectadas")
        st.info(f"Se detectaron anomalías en **{len(anomalies_df)}** productos. Los meses anómalos fueron **excluidos** del cálculo de compra.")
        st.dataframe(anomalies_df, use_container_width=True, hide_index=True)

    # --- 9. Tabla principal ---
    st.markdown("---")
    st.subheader("📑 Tabla Principal de Abastecimiento")

    f1, f2, f3, f4, f5 = st.columns(5)
    busq_codigo = f1.text_input("Buscar por Código")
    busq_desc = f2.text_input("Buscar por Descripción")
    filtro_familia = f3.multiselect("Familia", sorted(resultado_df["familia"].dropna().unique()))
    filtro_unidad = f4.multiselect("Unidad de Medida", sorted(resultado_df["unidad"].dropna().unique()))
    filtro_estado = f5.multiselect("Estado", sorted(resultado_df["estado"].unique()))

    tabla = resultado_df.copy()
    if busq_codigo:
        tabla = tabla[tabla["codigo"].str.contains(busq_codigo, case=False, na=False)]
    if busq_desc:
        tabla = tabla[tabla["descripcion"].str.contains(busq_desc, case=False, na=False)]
    if filtro_familia:
        tabla = tabla[tabla["familia"].isin(filtro_familia)]
    if filtro_unidad:
        tabla = tabla[tabla["unidad"].isin(filtro_unidad)]
    if filtro_estado:
        tabla = tabla[tabla["estado"].isin(filtro_estado)]

    orden_col = st.selectbox(
        "Ordenar por",
        ["Sin orden", "Consumo (mayor a menor)", "Días de cobertura (menor a mayor)", "Cantidad a comprar (mayor a menor)"],
    )
    if orden_col == "Consumo (mayor a menor)":
        tabla = tabla.sort_values("consumo_mensual", ascending=False)
    elif orden_col == "Días de cobertura (menor a mayor)":
        tabla = tabla.sort_values("dias_cobertura_display", ascending=True, na_position="last")
    elif orden_col == "Cantidad a comprar (mayor a menor)":
        tabla = tabla.sort_values("cantidad_comprar", ascending=False)

    # Tabla con comentarios
    cols_display = [
        "codigo", "descripcion", "familia", "unidad", 
        "stock_actual", "stock_origen",
        "consumo_mensual_historico", "consumo_mensual", "comentario_consumo",
        "consumo_diario", 
        "dias_cobertura_display", "comentario_dias_cobertura",
        "lead_time", 
        "stock_seguridad", "comentario_stock_seguridad",
        "punto_pedido", "comentario_punto_pedido",
        "stock_objetivo", "comentario_stock_objetivo",
        "cantidad_comprar", 
        "estado",
        "metodo_pronostico_usado", "mape_pronostico", "comentario_pronostico",
        "comentario_anomalias"
    ]
    # Filtrar columnas que existen
    cols_display = [col for col in cols_display if col in tabla.columns]
    tabla_display = tabla[cols_display].rename(columns={
        "codigo": "Código", 
        "descripcion": "Descripción", 
        "familia": "Familia", 
        "unidad": "U.M.",
        "stock_actual": "Stock Actual", 
        "stock_origen": "Origen",
        "consumo_mensual_historico": "Consumo Histórico", 
        "consumo_mensual": "Consumo (para compra)", 
        "comentario_consumo": "📝 Nota Consumo",
        "consumo_diario": "Consumo Diario",
        "dias_cobertura_display": "Días Cobertura", 
        "comentario_dias_cobertura": "📝 Nota Cobertura",
        "lead_time": "Lead Time", 
        "stock_seguridad": "Stock Seguridad",
        "comentario_stock_seguridad": "📝 Nota Seguridad",
        "punto_pedido": "Punto Pedido", 
        "comentario_punto_pedido": "📝 Nota Punto Pedido",
        "stock_objetivo": "Stock Objetivo", 
        "comentario_stock_objetivo": "📝 Nota Stock Objetivo",
        "cantidad_comprar": "Cantidad Comprar",
        "estado": "Estado",
        "metodo_pronostico_usado": "Método Pronóstico", 
        "mape_pronostico": "Error MAPE (%)", 
        "comentario_pronostico": "📝 Nota Pronóstico",
        "comentario_anomalias": "📝 Nota Anomalías"
    })
    st.dataframe(tabla_display.round(2), use_container_width=True, hide_index=True, height=500)

    # --- 11. Análisis por Familia ---
    st.markdown("---")
    st.subheader("🗂️ Análisis por Familia")
    familias_disponibles = sorted(resultado_df["familia"].dropna().unique())
    if familias_disponibles:
        familia_sel = st.selectbox("Selecciona una Familia", familias_disponibles)
        fam_df = resultado_df[resultado_df["familia"] == familia_sel]

        fc1, fc2, fc3, fc4 = st.columns(4)
        fc1.metric("Productos", f"{len(fam_df):,}")
        fc2.metric("Stock total", f"{fam_df['stock_actual'].sum():,.0f}")
        fc3.metric("Consumo mensual", f"{fam_df['consumo_mensual'].sum():,.0f}")
        fc4.metric("Compra sugerida", f"{fam_df['cantidad_comprar'].sum():,.0f}")

        fc5, fc6, fc7 = st.columns(3)
        fc5.metric("🔴 Críticos", f"{(fam_df['estado'] == '🔴 CRÍTICO').sum():,}")
        fc6.metric("🟡 Por abastecer", f"{(fam_df['estado'] == '🟡 POR ABASTECER').sum():,}")
        fc7.metric("🔵 Exceso", f"{(fam_df['estado'] == '🔵 EXCESO').sum():,}")

        st.markdown("**Principales productos por consumo:**")
        st.dataframe(
            fam_df.nlargest(10, "consumo_mensual")[["codigo", "descripcion", "unidad", "consumo_mensual", "estado"]]
            .rename(columns={"codigo": "Código", "descripcion": "Descripción", "unidad": "U.M.",
                              "consumo_mensual": "Consumo Mensual", "estado": "Estado"}),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("No hay Familias registradas en los archivos cargados.")

    # --- 12. Recomendación automática por producto ---
    st.markdown("---")
    st.subheader("💡 Recomendación Automática por Producto")
    producto_sel = st.selectbox(
        "Selecciona un producto",
        resultado_df["codigo"] + " - " + resultado_df["descripcion"].fillna(""),
    )
    codigo_sel = producto_sel.split(" - ")[0]
    prod = resultado_df[resultado_df["codigo"] == codigo_sel].iloc[0]

    pc1, pc2 = st.columns([1, 2])
    with pc1:
        st.markdown(f"**{prod['descripcion']} – {prod['codigo']}**")
        st.write(f"Familia: {prod['familia']}")
        st.write(f"Unidad: {prod['unidad']}")
        st.write(f"Stock actual: {prod['stock_actual']:.0f} {prod['unidad']}")
        st.write(f"Consumo histórico: {prod['consumo_mensual_historico']:.2f} {prod['unidad']}")
        st.write(f"Consumo (para compra): {prod['consumo_mensual']:.2f} {prod['unidad']}")
        with st.expander("📝 Ver detalle del consumo"):
            st.write(prod.get('comentario_consumo', 'No disponible'))
        st.write(f"Consumo diario: {prod['consumo_diario']:.2f} {prod['unidad']}")
        cobertura_txt = f"{prod['dias_cobertura']:.2f} días" if np.isfinite(prod["dias_cobertura"]) else "N/A (sin consumo)"
        st.write(f"Cobertura: {cobertura_txt}")
        with st.expander("📝 Ver detalle de cobertura"):
            st.write(prod.get('comentario_dias_cobertura', 'No disponible'))
        st.write(f"Lead Time: {prod['lead_time']:.0f} días")
        st.write(f"Stock seguridad: {prod['stock_seguridad']:.0f} {prod['unidad']}")
        with st.expander("📝 Ver detalle de stock seguridad"):
            st.write(prod.get('comentario_stock_seguridad', 'No disponible'))
        st.write(f"Punto de pedido: {prod['punto_pedido']:.0f} {prod['unidad']}")
        with st.expander("📝 Ver detalle de punto de pedido"):
            st.write(prod.get('comentario_punto_pedido', 'No disponible'))
        st.write(f"Stock objetivo: {prod['stock_objetivo']:.0f} {prod['unidad']}")
        with st.expander("📝 Ver detalle de stock objetivo"):
            st.write(prod.get('comentario_stock_objetivo', 'No disponible'))
        st.write(f"Compra recomendada: {prod['cantidad_comprar']:.0f} {prod['unidad']}")
        st.write(f"Estado: {prod['estado']}  |  Prioridad: {prod['prioridad_compra']}")
        with st.expander("📝 Ver método de pronóstico"):
            st.write(prod.get('comentario_pronostico', 'No disponible'))
        if prod.get('comentario_anomalias', 'Sin anomalías detectadas') != "Sin anomalías detectadas":
            with st.expander("📝 Ver anomalías detectadas"):
                st.write(prod.get('comentario_anomalias', 'Sin anomalías detectadas'))
    with pc2:
        st.info(prod["recomendacion"])
        prod_monthly = monthly_df[monthly_df["codigo"] == codigo_sel].sort_values("periodo_str")
        if not prod_monthly.empty:
            fig = px.line(prod_monthly, x="periodo_str", y="consumo", markers=True,
                           title="Comportamiento histórico del producto",
                           labels={"periodo_str": "Mes", "consumo": "Consumo"})
            st.plotly_chart(fig, use_container_width=True)

    # --- 13. Priorización de compras ---
    st.markdown("---")
    st.subheader("🏷️ Priorización de Compras")
    prioridad_tabla = resultado_df[resultado_df["cantidad_comprar"] > 0][
        ["codigo", "descripcion", "familia", "unidad", "cantidad_comprar", "prioridad_compra", "estado"]
    ].sort_values(["prioridad_compra", "cantidad_comprar"], ascending=[True, False]).rename(columns={
        "codigo": "Código", "descripcion": "Descripción", "familia": "Familia", "unidad": "U.M.",
        "cantidad_comprar": "Cantidad a Comprar", "prioridad_compra": "Prioridad de Compra", "estado": "Estado",
    })
    st.dataframe(prioridad_tabla, use_container_width=True, hide_index=True, height=350)

    # --- 14. Exportación ---
    st.markdown("---")
    st.subheader("⬇️ 5️⃣ Descargar Resultado en Excel")
    if st.button("Generar archivo Excel"):
        with st.spinner("Generando archivo Excel con todas las hojas..."):
            output_path = "/tmp/resultado_abastecimiento.xlsx"
            build_export_excel(
                resultado_df, 
                monthly_df, 
                issues, 
                params, 
                anomalies_df, 
                products_added_df, 
                output_path
            )
            with open(output_path, "rb") as f:
                excel_bytes = f.read()
        st.download_button(
            label="📥 Descargar Excel de Resultado",
            data=excel_bytes,
            file_name=f"resultado_abastecimiento_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Sube ambos archivos Excel y presiona **CALCULAR ABASTECIMIENTO** para ver el análisis completo.")
