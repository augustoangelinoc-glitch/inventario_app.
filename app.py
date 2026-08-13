# -*- coding: utf-8 -*-
"""Genera las plantillas Excel y los datos de prueba con anomalías."""
import os
import random
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

random.seed(42)
np.random.seed(42)

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(OUT_DIR, "datos_ejemplo")
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# 1) PLANTILLAS
# ---------------------------------------------------------------------------
plantilla_stock = pd.DataFrame(
    {
        "Código": ["P001", "P002", "P003"],
        "Descripción": ["Filtro de aceite", "Aceite hidráulico", "Manguera hidráulica"],
        "Familia": ["Filtros", "Lubricantes", "Mangueras"],
        "Unidad de Medida": ["UND", "LT", "MT"],
        "Stock Actual": [25, 150, 80],
    }
)
plantilla_stock.to_excel(os.path.join(DATA_DIR, "Plantilla_Stock_Actual.xlsx"), index=False)

plantilla_salidas = pd.DataFrame(
    {
        "Código": ["P001", "P001", "P002"],
        "Descripción": ["Filtro de aceite", "Filtro de aceite", "Aceite hidráulico"],
        "Familia": ["Filtros", "Filtros", "Lubricantes"],
        "Unidad de Medida": ["UND", "UND", "LT"],
        "Fecha": ["05/01/2026", "15/02/2026", "12/03/2026"],
        "Cantidad Salida": [5, 8, 30],
    }
)
plantilla_salidas.to_excel(os.path.join(DATA_DIR, "Plantilla_Salidas_Historicas.xlsx"), index=False)

# Plantillas de parámetros
plantilla_params_familia = pd.DataFrame(
    {
        "Familia": ["Filtros", "Lubricantes", "Mangueras"],
        "Lead Time (días)": [10, 5, 7],
        "Días Cobertura Objetivo": [20, 10, 15],
        "% Stock Seguridad": [25, 15, 20],
        "Método Cálculo": ["Últimos N meses", "Todo el histórico", "Promedio ponderado"],
        "N° Meses": [6, 3, 4]
    }
)
plantilla_params_familia.to_excel(
    os.path.join(DATA_DIR, "Plantilla_Parametros_Familia.xlsx"), index=False
)

plantilla_params_producto = pd.DataFrame(
    {
        "Código": ["P001", "P002"],
        "Lead Time (días)": [14, 3],
        "Días Cobertura Objetivo": [30, 7],
        "% Stock Seguridad": [30, 10]
    }
)
plantilla_params_producto.to_excel(
    os.path.join(DATA_DIR, "Plantilla_Parametros_Producto.xlsx"), index=False
)

# ---------------------------------------------------------------------------
# 2) DATOS DE PRUEBA REALISTAS CON ANOMALÍAS
# ---------------------------------------------------------------------------
familias_catalogo = {
    "Filtros": ["UND"], "Lubricantes": ["LT", "GAL"], "Mangueras": ["MT"],
    "Repuestos Mecánicos": ["UND", "JGO"], "Pernos y Tuercas": ["UND", "CAJA"],
    "Correas": ["UND"], "Rodamientos": ["UND", "JGO"], "Empaquetaduras": ["UND", "JGO"],
    "Neumáticos": ["UND"], "Baterías": ["UND"], "Herramientas": ["UND", "JGO"],
    "Pinturas y Recubrimientos": ["GAL", "LT"], "Equipos de Seguridad": ["UND", "PAR"],
    "Eléctricos": ["UND", "MT"], "Soldadura": ["UND", "KG"], "Grasas y Aditivos": ["KG", "LT"],
    "Válvulas": ["UND"], "Sensores": ["UND"],
}

descripciones_base = [
    "Filtro de aceite", "Filtro de aire", "Filtro de combustible", "Aceite hidráulico",
    "Aceite de motor 15W40", "Grasa multipropósito", "Manguera hidráulica 1/2\"",
    "Manguera hidráulica 3/4\"", "Correa en V", "Correa dentada", "Rodamiento cónico",
    "Rodamiento de bolas", "Empaquetadura de culata", "Perno hexagonal 1/2\"",
    "Tuerca hexagonal 1/2\"", "Neumático 295/80 R22.5", "Batería 12V 150Ah",
    "Llave mixta 1/2\"", "Pintura anticorrosiva", "Guantes de seguridad",
    "Casco de seguridad", "Cable eléctrico THW", "Electrodo de soldadura",
    "Grasa para rodamientos", "Válvula de alivio", "Sensor de presión",
    "Terminal hidráulico", "Bomba hidráulica", "Retén de eje", "Bujía de encendido",
]

n_productos = 120
codigos = [f"P{str(i).zfill(4)}" for i in range(1, n_productos + 1)]

stock_rows = []
producto_info = {}
for i, codigo in enumerate(codigos):
    familia = random.choice(list(familias_catalogo.keys()))
    unidad = random.choice(familias_catalogo[familia])
    desc = f"{random.choice(descripciones_base)} - Ref {i+1:03d}"
    stock_actual = int(np.random.gamma(shape=2.0, scale=40))
    producto_info[codigo] = {"familia": familia, "unidad": unidad, "descripcion": desc}
    stock_rows.append(
        {
            "Código": codigo, "Descripción": desc, "Familia": familia,
            "Unidad de Medida": unidad, "Stock Actual": stock_actual,
        }
    )

stock_df = pd.DataFrame(stock_rows)

# Introducir inconsistencias intencionales
stock_df.loc[5, "Descripción"] = None            # código sin descripción
stock_df.loc[10, "Familia"] = None               # producto sin familia
stock_df.loc[15, "Unidad de Medida"] = None       # producto sin unidad
dup_row = stock_df.iloc[20].copy()
stock_df = pd.concat([stock_df, pd.DataFrame([dup_row])], ignore_index=True)  # código duplicado
stock_df.loc[25, "Stock Actual"] = None           # stock vacío
# Stock negativo
stock_df.loc[30, "Stock Actual"] = -5            # stock negativo

stock_df.to_excel(os.path.join(DATA_DIR, "Stock_Actual_PRUEBA.xlsx"), index=False)

# --- Salidas históricas: enero 2026 a julio 2026 ---
meses = pd.date_range("2026-01-01", "2026-07-01", freq="MS")
salidas_rows = []
for codigo, info in producto_info.items():
    consumo_base = max(np.random.gamma(shape=1.5, scale=15), 1)
    tendencia = random.choice([-1, 0, 1])
    for m_idx, mes in enumerate(meses):
        factor_tendencia = 1 + (tendencia * 0.08 * m_idx)
        consumo_mes = max(consumo_base * factor_tendencia * np.random.uniform(0.7, 1.3), 0)
        n_movs = random.randint(1, 4)
        cantidades = np.random.dirichlet(np.ones(n_movs)) * consumo_mes
        for cant in cantidades:
            dia = random.randint(1, 27)
            fecha = mes + timedelta(days=dia - 1)
            if fecha > datetime(2026, 7, 31):
                continue
            salidas_rows.append(
                {
                    "Código": codigo,
                    "Descripción": info["descripcion"],
                    "Familia": info["familia"],
                    "Unidad de Medida": info["unidad"],
                    "Fecha": fecha.strftime("%d/%m/%Y"),
                    "Cantidad Salida": round(float(cant), 1),
                }
            )

# Producto con una sola salida en todo el histórico (anomalía)
salidas_rows.append({
    "Código": "P999",
    "Descripción": "Producto con una sola salida",
    "Familia": "Repuestos Mecánicos",
    "Unidad de Medida": "UND",
    "Fecha": "15/03/2026",
    "Cantidad Salida": 7.0,
})

# Producto con stock 0 pero consumo constante (rotura de stock)
for mes in meses:
    salidas_rows.append({
        "Código": "P888",
        "Descripción": "Producto en rotura de stock",
        "Familia": "Repuestos Mecánicos",
        "Unidad de Medida": "UND",
        "Fecha": mes.strftime("%d/%m/%Y"),
        "Cantidad Salida": 10.0,
    })

salidas_df = pd.DataFrame(salidas_rows)
salidas_df = salidas_df[~salidas_df["Código"].isin(["P0002", "P0004"])]

# Productos con salidas que no están en stock
extra_salidas = pd.DataFrame(
    {
        "Código": ["P9001", "P9001"],
        "Descripción": ["Producto descontinuado", "Producto descontinuado"],
        "Familia": ["Repuestos Mecánicos", "Repuestos Mecánicos"],
        "Unidad de Medida": ["UND", "UND"],
        "Fecha": ["10/02/2026", "10/04/2026"],
        "Cantidad Salida": [3, 2],
    }
)
salidas_df = pd.concat([salidas_df, extra_salidas], ignore_index=True)

# Fecha incorrecta y cantidad inválida
salidas_df.loc[salidas_df.index[3], "Fecha"] = "31/13/2026"      # fecha inválida
salidas_df.loc[salidas_df.index[7], "Cantidad Salida"] = -5      # cantidad inválida

salidas_df.to_excel(os.path.join(DATA_DIR, "Salidas_Historicas_PRUEBA.xlsx"), index=False)

print(f"✅ Productos en stock de prueba: {len(stock_df)}")
print(f"✅ Registros de salidas de prueba: {len(salidas_df)}")
print(f"✅ Archivos generados en: {DATA_DIR}")