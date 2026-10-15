# 🏷️ E-Commerce Price & Stock Tracker with Analytics Export

Herramienta en Python para monitorear precios de un catálogo e-commerce, guardar historial y generar reportes en Excel en **dólares estadounidenses (USD)**, pensada para operar en EE.UU. y Ecuador (ambos usan el dólar).

## 🚀 Características
- **Extracción de precios y stock** con paginación, timeout y manejo de errores.
- **Conversión GBP → USD** con tipo de cambio configurable (`GBP_USD_RATE`).
- **Historial de precios** (`historial_precios.csv`): cada corrida se guarda y se calcula la variación vs la anterior.
- **Excel con formato `$`**, encabezados, filtros, paneles congelados y formato condicional (ahorros altos en verde, alzas de precio en rojo).
- **Hoja "Resumen"** con fecha, fuente, tipo de cambio y promedios.

## ⚠️ Alcance y limitaciones
- La fuente es [books.toscrape.com](http://books.toscrape.com), un sitio de práctica para scraping.
- El **precio de lista es simulado** (markup 10%–20%, estable por producto) porque el sitio no publica precios de lista reales.
- Los precios no incluyen impuestos (IVA / sales tax), aranceles ni envío.

## 🛠️ Tech Stack
Python 3.10+ · Requests · BeautifulSoup4 · Pandas · OpenPyXL

## 📦 Instalación y uso

```bash
git clone https://github.com/Martin-Enriquez/ecommerce-price-tracker-analytics.git
cd ecommerce-price-tracker-analytics
pip install -r requirements.txt

# Opcional: actualizar el tipo de cambio GBP -> USD
GBP_USD_RATE=1.34 python main.py
```

En Windows (PowerShell):

```powershell
$env:GBP_USD_RATE="1.34"; python main.py
```

## 📁 Salidas
| Archivo | Contenido |
|---|---|
| `reporte_monitor_precios.xlsx` | Reporte con hojas "Monitor de Precios" y "Resumen" |
| `historial_precios.csv` | Snapshots de cada corrida (precio GBP, tipo de cambio, precio USD, stock) |
