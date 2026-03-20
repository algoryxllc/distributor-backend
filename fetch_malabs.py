import requests
from openpyxl import Workbook
import os

EMAIL = os.environ.get("MALABS_EMAIL")
PASSWORD = os.environ.get("MALABS_PASSWORD")

wb = Workbook()
ws = wb.active
ws.title = "MA Labs Catalog"

# Headers
ws.append([
    "List No", "Item No", "UPC Code", "Manufacturer No",
    "Manufacturer", "Category", "Product Name", "Price",
    "Instant Rebate", "Weight", "Length", "Width", "Height",
    "Package", "Is Domestic Only",
    "Warehouse 1001", "Warehouse 1002", "Warehouse 1003",
    "Warehouse 1004", "Warehouse 1005", "Warehouse 1006"
])

print("Connecting to MA Labs API...")

# Step 1: Get accurate total from API (no filters)
first = requests.get(
    "https://online.malabs.com/mws/items/?format=json&page=1",
    auth=(EMAIL, PASSWORD),
    timeout=30
)

data = first.json()
total_items = data.get("count", 0)
total_pages = (total_items // 10) + 1

print(f"Total items in MA Labs catalog: {total_items}")
print(f"Total pages to fetch: {total_pages}")
print("Starting full catalog fetch...\n")

seen = set()  # Track seen item numbers to avoid duplicates
page = 1
total_fetched = 0

while True:
    url = f"https://online.malabs.com/mws/items/?format=json&page={page}"
    response = requests.get(url, auth=(EMAIL, PASSWORD), timeout=30)

    if response.status_code != 200:
        print(f"Error on page {page}: {response.status_code}")
        break

    page_data = response.json()
    results = page_data.get("results", [])

    if not results:
        print("No more results found. Stopping.")
        break

    for item in results:
        item_no = item.get("item_no", "")

        # Skip duplicates
        if item_no in seen:
            continue
        seen.add(item_no)

        inventory = item.get("inventory", {})

        ws.append([
            item.get("list_no", ""),
            item_no,
            item.get("upc_code", ""),
            item.get("manufacturer_no", ""),
            item.get("manufacturer", ""),
            item.get("category", ""),
            item.get("product_name", ""),
            item.get("price", ""),
            item.get("instant_rebate", ""),
            item.get("weight", ""),
            item.get("length", ""),
            item.get("width", ""),
            item.get("height", ""),
            item.get("package", ""),
            item.get("is_domestic_only", ""),
            inventory.get("1001", 0),
            inventory.get("1002", 0),
            inventory.get("1003", 0),
            inventory.get("1004", 0),
            inventory.get("1005", 0),
            inventory.get("1006", 0),
        ])
        total_fetched += 1

    print(f"Page {page}/{total_pages} — {total_fetched} unique items so far...")

    # Stop when all pages fetched
    if page >= total_pages:
        print("All pages fetched!")
        break

    page += 1

wb.save("malabs_catalog.xlsx")
print(f"\nDone! Total {total_fetched} unique items saved to malabs_catalog.xlsx")
