import requests
from openpyxl import Workbook
import os

TOKEN = os.environ.get("VOIP_TOKEN")

def fetch_all_products():
    print("Fetching all 888VoIP products...")
    response = requests.get(
        "https://stagingapi.888voip.com/api/products",
        headers={
            "Authorization": f"Bearer {TOKEN}",
            "Accept": "application/json"
        },
        timeout=60
    )
    print(f"Products response: {response.status_code}")
    if response.status_code == 200:
        return response.json().get("products", [])
    else:
        print(f"Error: {response.status_code} - {response.text}")
        return []

# Fetch all products
products = fetch_all_products()
print(f"Total products fetched: {len(products)}")

if not products:
    print("No products found. Exiting.")
    exit(1)

# Create Excel
wb = Workbook()
ws = wb.active
ws.title = "888VoIP Catalog"

# Headers exactly as per document
ws.append([
    "sku",
    "partNumber",
    "name",
    "description",
    "price",
    "msrp",
    "make",
    "weight",
    "length",
    "width",
    "height",
    "qty",
    "categories",
    "stock_BUF",
    "stock_RNO"
])

for product in products:
    stock = product.get("stockByWarehouse", {})
    ws.append([
        product.get("sku", ""),
        product.get("partNumber", ""),
        product.get("name", ""),
        product.get("description", ""),
        product.get("price", ""),
        product.get("msrp", ""),
        product.get("make", ""),
        product.get("weight", ""),
        product.get("length", ""),
        product.get("width", ""),
        product.get("height", ""),
        product.get("qty", ""),
        product.get("categories", ""),
        stock.get("BUF", 0),
        stock.get("RNO", 0)
    ])

wb.save("voip_catalog.xlsx")
print(f"\nDone! {len(products)} products saved to voip_catalog.xlsx")
