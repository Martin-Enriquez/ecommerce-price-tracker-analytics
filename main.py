"""
Ecommerce Price & Stock Tracker with Conditional Formatting Excel Export
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import re

# URL de ejemplo de e-commerce de prueba
CATALOG_URL = "http://books.toscrape.com/catalogue/category/books/travel_2/index.html"

def fetch_products():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    response = requests.get(CATALOG_URL, headers=headers)
    if response.status_code != 200:
        print(f"Error al conectar con la tienda: {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    products = []

    for item in soup.select("article.product_pod"):
        title = item.h3.a["title"]
        price_text = item.select_one("p.price_color").text
        price = float(re.search(r"\d+\.\d+", price_text).group())
        availability = item.select_one("p.instock.availability").text.strip()
        
        # Simulación de Precio Anterior / Lista para calcular descuento
        # (Simulamos un precio de lista un 10%-20% mayor para demostrar análisis comercial)
        list_price = round(price * 1.15, 2)
        discount_pct = round(((list_price - price) / list_price) * 100, 1)

        products.append({
            "Producto": title,
            "Precio Lista (£)": list_price,
            "Precio Oferta (£)": price,
            "Ahorro (%)": discount_pct,
            "Estado Stock": availability
        })

    return products

def export_price_analytics(data, filename="reporte_monitor_precios.xlsx"):
    df = pd.DataFrame(data)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Monitor de Precios")
        ws = writer.sheets["Monitor de Precios"]

        # Estilos corporativos
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center = Alignment(horizontal="center", vertical="center")

        # Formato de Encabezados
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center

        # Formato de celdas y números
        price_list_col = 2
        price_offer_col = 3
        savings_col = 4

        for row in ws.iter_rows(min_row=2):
            row[price_list_col - 1].number_format = '"£"#,##0.00'
            row[price_offer_col - 1].number_format = '"£"#,##0.00'
            row[savings_col - 1].number_format = '0.0"%"'
            row[savings_col - 1].alignment = center
            row[4].alignment = center  # Stock

        # Regla de Formato Condicional (Verde para descuentos mayores al 10%)
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font = Font(color="006100", bold=True)
        rule = CellIsRule(operator="greaterThan", formula=["10"], fill=green_fill, font=green_font)
        ws.conditional_formatting.add(f"D2:D{len(df)+1}", rule)

        # Ajuste dinámico de ancho de columna
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    print(f"✅ Reporte de monitoreo exportado exitosamente a {filename}")

if __name__ == "__main__":
    print("Iniciando monitor de precios...")
    data = fetch_products()
    if data:
        export_price_analytics(data)