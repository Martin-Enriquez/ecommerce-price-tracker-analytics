"""
E-commerce Price & Stock Tracker con historial y exportación a Excel (USD).

- Extrae precios del catálogo (sitio de práctica books.toscrape.com, precios en GBP).
- Convierte a USD con un tipo de cambio configurable (GBP_USD_RATE).
- Guarda cada corrida en un historial CSV y calcula la variación vs la corrida anterior.
- Exporta un Excel con formato "$" (USD), formato condicional y hoja de resumen.

NOTA: el "Precio Lista" es SIMULADO (markup aleatorio 10%-20%, estable por producto),
porque el sitio de práctica no publica precios de lista reales.
"""

import os
import random
import re
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------- Configuración -----------------------------
CATALOG_URL = "http://books.toscrape.com/catalogue/category/books/travel_2/index.html"
OUTPUT_FILE = "reporte_monitor_precios.xlsx"
HISTORY_FILE = "historial_precios.csv"

# Tipo de cambio GBP -> USD. Actualízalo antes de cada corrida real:
#   Linux/Mac: GBP_USD_RATE=1.34 python main.py
GBP_USD_RATE = float(os.getenv("GBP_USD_RATE", "1.34"))

SIM_MARKUP_RANGE = (0.10, 0.20)   # markup simulado para el precio de lista
HIGHLIGHT_DISCOUNT = 0.13         # >= 13% de ahorro se resalta en verde
REQUEST_TIMEOUT = 15              # segundos
USD_FORMAT = '"$"#,##0.00'        # muestra "$" sin depender de la configuración regional
PCT_FORMAT = "0.0%"

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


# ------------------------------ Extracción -------------------------------
def fetch_products():
    """Recorre el catálogo (con paginación) y devuelve una lista de productos."""
    session = requests.Session()
    session.headers.update(HEADERS)
    products, url = [], CATALOG_URL

    while url:
        try:
            response = session.get(url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
        except requests.RequestException as exc:
            print(f"❌ Error al conectar con la tienda ({url}): {exc}")
            break

        response.encoding = "utf-8"  # evita caracteres rotos (â\x80\x99, Â£)
        soup = BeautifulSoup(response.text, "html.parser")

        for item in soup.select("article.product_pod"):
            title = item.h3.a["title"]
            price_text = item.select_one("p.price_color").text
            price_gbp = float(re.search(r"\d+(?:\.\d+)?", price_text).group())
            availability = item.select_one("p.instock.availability").text.strip()
            products.append(
                {"title": title, "price_gbp": price_gbp, "availability": availability}
            )

        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None

    return products


# ------------------------------- Análisis --------------------------------
def simulated_list_price(title, offer_price_usd):
    """Precio de lista simulado: markup 10%-20% estable por producto (no cambia entre corridas)."""
    markup = random.Random(title).uniform(*SIM_MARKUP_RANGE)
    return round(offer_price_usd * (1 + markup), 2)


def build_dataframe(products, run_ts):
    rows = []
    for p in products:
        offer_usd = round(p["price_gbp"] * GBP_USD_RATE, 2)
        list_usd = simulated_list_price(p["title"], offer_usd)
        rows.append(
            {
                "Producto": p["title"],
                "Precio Lista Simulado ($)": list_usd,
                "Precio Oferta ($)": offer_usd,
                "Ahorro ($)": round(list_usd - offer_usd, 2),
                "Ahorro (%)": round((list_usd - offer_usd) / list_usd, 4),  # fracción: 0.13 = 13%
                "Estado Stock": p["availability"],
                "_price_gbp": p["price_gbp"],
            }
        )
    df = pd.DataFrame(rows)
    df["Var. vs Corrida Anterior (%)"] = None
    return df


def apply_history(df, run_ts):
    """Compara contra la última corrida guardada y agrega la corrida actual al historial."""
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        last = hist.sort_values("timestamp").groupby("Producto").tail(1).set_index("Producto")
        prev_gbp = df["Producto"].map(last["precio_gbp"])
        df["Var. vs Corrida Anterior (%)"] = ((df["_price_gbp"] - prev_gbp) / prev_gbp).round(4)

    snapshot = pd.DataFrame(
        {
            "timestamp": run_ts,
            "Producto": df["Producto"],
            "precio_gbp": df["_price_gbp"],
            "tipo_cambio": GBP_USD_RATE,
            "precio_oferta_usd": df["Precio Oferta ($)"],
            "stock": df["Estado Stock"],
        }
    )
    snapshot.to_csv(HISTORY_FILE, mode="a", header=not os.path.exists(HISTORY_FILE),
                    index=False, encoding="utf-8")
    return df.drop(columns=["_price_gbp"])


# ------------------------------- Exportación -----------------------------
def export_price_analytics(df, run_ts, filename=OUTPUT_FILE):
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Monitor de Precios")
        ws = writer.sheets["Monitor de Precios"]
        col = {name: i + 1 for i, name in enumerate(df.columns)}  # nombre -> nº de columna

        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center = Alignment(horizontal="center", vertical="center")
        for cell in ws[1]:
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        usd_cols = ["Precio Lista Simulado ($)", "Precio Oferta ($)", "Ahorro ($)"]
        pct_cols = ["Ahorro (%)", "Var. vs Corrida Anterior (%)"]
        for row in ws.iter_rows(min_row=2):
            for name in usd_cols:
                row[col[name] - 1].number_format = USD_FORMAT
            for name in pct_cols:
                row[col[name] - 1].number_format = PCT_FORMAT
                row[col[name] - 1].alignment = center
            row[col["Estado Stock"] - 1].alignment = center

        # Verde: ahorro >= umbral
        letter = get_column_letter(col["Ahorro (%)"])
        rule = CellIsRule(
            operator="greaterThanOrEqual", formula=[str(HIGHLIGHT_DISCOUNT)],
            fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
            font=Font(color="006100", bold=True),
        )
        ws.conditional_formatting.add(f"{letter}2:{letter}{len(df) + 1}", rule)

        # Rojo: el precio subió vs la corrida anterior
        letter = get_column_letter(col["Var. vs Corrida Anterior (%)"])
        up = CellIsRule(
            operator="greaterThan", formula=["0"],
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            font=Font(color="9C0006", bold=True),
        )
        ws.conditional_formatting.add(f"{letter}2:{letter}{len(df) + 1}", up)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for c in ws.columns:
            max_len = max(len(str(x.value or "")) for x in c)
            ws.column_dimensions[get_column_letter(c[0].column)].width = min(max(max_len + 4, 14), 60)

        # Hoja de resumen / metadatos
        summary = pd.DataFrame(
            {
                "Concepto": [
                    "Fecha de corrida", "Fuente", "Moneda de origen", "Moneda del reporte",
                    "Tipo de cambio GBP→USD", "Productos monitoreados",
                    "Precio oferta promedio ($)", "Ahorro promedio (%)", "Nota",
                ],
                "Valor": [
                    run_ts, CATALOG_URL, "GBP (£)", "USD ($)", GBP_USD_RATE, len(df),
                    round(df["Precio Oferta ($)"].mean(), 2), round(df["Ahorro (%)"].mean(), 4),
                    "El precio de lista es simulado (markup 10%-20%). Precios sin impuestos ni envío.",
                ],
            }
        )
        summary.to_excel(writer, index=False, sheet_name="Resumen")
        ws2 = writer.sheets["Resumen"]
        ws2["B8"].number_format = USD_FORMAT
        ws2["B9"].number_format = PCT_FORMAT
        for cell in ws2[1]:
            cell.fill, cell.font = header_fill, header_font
        ws2.column_dimensions["A"].width = 30
        ws2.column_dimensions["B"].width = 80
        for r in ws2.iter_rows(min_row=2):
            r[1].alignment = Alignment(horizontal="left", wrap_text=True)

    print(f"✅ Reporte exportado a {filename}")


if __name__ == "__main__":
    print("Iniciando monitor de precios...")
    products = fetch_products()
    if products:
        run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        frame = apply_history(build_dataframe(products, run_ts), run_ts)
        export_price_analytics(frame, run_ts)
    else:
        print("No se obtuvieron productos; no se generó reporte.")
