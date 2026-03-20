import requests
from openpyxl import Workbook
import os

EMAIL = os.environ.get("MALABS_EMAIL")
PASSWORD = os.environ.get("MALABS_PASSWORD")

wb = Workbook()
ws = wb.active
ws.title = "MA Labs Catalog"

ws.append(["Item No", "UPC Code", "Manufacturer No", "Manufacturer",
           "Category", "Product Name", "Price", "Inventory"])

print("Connecting to MA Labs API...")

# Get total count
first = requests.get(
    "https://online.malabs.com/mws/items/?format=json&page=1",
    auth=(EMAIL, PASSWORD),
    timeout=30
)

data = first.json()
total_items = data.get("count", 0)
total_pages = (total_items // 10) + 1

print(f"Total items: {total_items} | Total pages: {total_pages}")

page = 1
total_fetched = 0

while True:
    url = f"https://online.malabs.com/mws/items/?format=json&page={page}"
    response = requests.get(url, auth=(EMAIL, PASSWORD), timeout=30)

    if response.status_code != 200:
        print(f"Error on page {page}: {response.status_code}")
        break

    results = response.json().get("results", [])

    if not results:
        break

    for item in results:
        inventory = item.get("inventory", {})
        total_inv = sum(inventory.values()) if inventory else 0
        ws.append([
            item.get("item_no", ""),
            item.get("upc_code", ""),
            item.get("manufacturer_no", ""),
            item.get("manufacturer", ""),
            item.get("category", ""),
            item.get("product_name", ""),
            item.get("price", ""),
            total_inv
        ])

    total_fetched += len(results)
    print(f"Page {page}/{total_pages} — {total_fetched}/{total_items} items fetched...")

    if page >= total_pages:
        break

    page += 1

wb.save("malabs_catalog.xlsx")
print(f"\nDone! Total {total_fetched} items saved to malabs_catalog.xlsx")
