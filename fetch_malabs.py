import requests
from openpyxl import Workbook
import os

EMAIL = os.environ.get("MALABS_EMAIL")
PASSWORD = os.environ.get("MALABS_PASSWORD")

wb = Workbook()
ws = wb.active
ws.title = "MA Labs Catalog"

ws.append([
    "list_no", "item_no", "upc_code", "manufacturer_no",
    "manufacturer", "category", "product_name", "price",
    "instant_rebate", "instant_rebate_item_no",
    "weight", "length", "width", "height",
    "package", "specorder", "is_domestic_only",
    "inventory_1001", "inventory_1002", "inventory_1003",
    "inventory_1004", "inventory_1005", "inventory_1006"
])

print("Connecting to MA Labs API...")

# Get total count
first = requests.get(
    "https://online.malabs.com/mws/items/?format=json&limit=10&offset=0",
    auth=(EMAIL, PASSWORD),
    timeout=30
)

data = first.json()
total_items = data.get("count", 0)
print(f"Total items: {total_items}")
print("Starting full catalog fetch...\n")

offset = 0
total_fetched = 0
next_url = f"https://online.malabs.com/mws/items/?format=json&limit=10&offset=0"

while next_url:
    try:
        response = requests.get(next_url, auth=(EMAIL, PASSWORD), timeout=60)
    except Exception as e:
        print(f"Timeout at offset {offset}, retrying... {e}")
        continue

    if response.status_code != 200:
        print(f"Error at offset {offset}: {response.status_code}")
        break

    page_data = response.json()
    results = page_data.get("results", [])

    if not results:
        print("No more results. Stopping.")
        break

    for item in results:
        inventory = item.get("inventory", {})
        ws.append([
            item.get("list_no", ""),
            item.get("item_no", ""),
            item.get("upc_code", ""),
            item.get("manufacturer_no", ""),
            item.get("manufacturer", ""),
            item.get("category", ""),
            item.get("product_name", ""),
            item.get("price", ""),
            item.get("instant_rebate", ""),
            item.get("instant_rebate_item_no", ""),
            item.get("weight", ""),
            item.get("length", ""),
            item.get("width", ""),
            item.get("height", ""),
            item.get("package", ""),
            item.get("specorder", ""),
            item.get("is_domestic_only", ""),
            inventory.get("1001", 0),
            inventory.get("1002", 0),
            inventory.get("1003", 0),
            inventory.get("1004", 0),
            inventory.get("1005", 0),
            inventory.get("1006", 0),
        ])
        total_fetched += 1

    offset += 10
    print(f"Offset {offset}/{total_items} — {total_fetched} items fetched...")

    # Use next URL from API response
    next_url = page_data.get("next")
    if next_url and "format=json" not in next_url:
        next_url = next_url + "&format=json"

wb.save("malabs_catalog.xlsx")
print(f"\nDone! {total_fetched} items saved to malabs_catalog.xlsx")
