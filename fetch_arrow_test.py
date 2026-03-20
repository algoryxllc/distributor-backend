import requests
import os
from openpyxl import Workbook

# =============================================
# PASTE YOUR ACCESS TOKEN HERE (from Postman)
ACCESS_TOKEN = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzY29wZSI6WyJNeUFycm93RXh0ZXJuYWxDbGllbnQiXSwiZXhwIjoxNzczOTk0NTI2LCJqdGkiOiJBMVJUZTA5b19pN3dXWXpsQjZIdkRxXzBjNGMiLCJjbGllbnRfaWQiOiJhbGdvcnl4LWxsYyJ9.NMfI-29IM546uF23OIckdUn3AiLSyoNbcwQTQAljxFK-4yVUIuVR3ejwhcPSfRz8DW9rVNej6JYMqkl6za1XsSkMykRizsXY7i2OsAHEUOuKcVwR_vP6Ht32ACwKaJDXRDiHJMxlSj2ae57s9KiCFYuj8G97pyoHkBKa8AJxn8jW8mJz3oO3CZBZpKDIByHtLoUUwKffmy19yKTUjcEUkXO_o7Rnz4XWYGa3PfUhy_q5RjcpXqe_nYSz7cJNeLheSRqIj6dduQk-UmOgtOtD-b9qPOG4mwMTp8r8z21SfFkCZ1-xUnx-gZDl9C4c7xaHRolbq-_Z4aMr-_ZalNOkUg"
# =============================================

CLIENT_ID = os.environ.get("ARROW_CLIENT_ID")
SKUS = os.environ.get("SKUS", "BAV99\nCT240BX200SSD1")

def fetch_part(token, sku):
    url = "https://my.arrow.com/api/priceandavail/parts"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "client_id": CLIENT_ID
    }
    params = {
        "search": sku,
        "currency": "USD",
        "quantity": 1,
        "pageNumber": 1,
        "pageSize": 200,
        "version": 3
    }
    response = requests.get(url, headers=headers, params=params, timeout=30)
    print(f"Fetch status for {sku}: {response.status_code}")
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: {response.text}")
        return None

# Parse SKUs
sku_list = [s.strip() for s in SKUS.strip().split("\n") if s.strip()]
print(f"Total SKUs: {len(sku_list)}")

# Create Excel
wb = Workbook()
ws = wb.active
ws.title = "Arrow Catalog"
ws.append([
    "Search SKU", "Part Number", "Manufacturer", "Description",
    "Price (USD)", "Inventory", "Min Order Qty", "Lead Time (weeks)",
    "Warehouse Code"
])

total_fetched = 0

for i, sku in enumerate(sku_list, 1):
    print(f"Fetching {i}/{len(sku_list)}: {sku}")
    data = fetch_part(ACCESS_TOKEN, sku)

    if not data:
        ws.append([sku, "ERROR", "", "", "", "", "", "", ""])
        continue

    pricing = data.get("pricingResponse", [])
    if not pricing:
        ws.append([sku, "NOT FOUND", "", "", "", "", "", "", ""])
        continue

    for item in pricing:
        pricing_tiers = item.get("pricingTier", [])
        price = pricing_tiers[0].get("resalePrice", "") if pricing_tiers else ""
        lead_time_obj = item.get("leadTime", {})
        lead_time = lead_time_obj.get("arrowLeadTime", "") if lead_time_obj else ""

        ws.append([
            sku,
            item.get("partNumber", ""),
            item.get("manufacturer", ""),
            item.get("description", ""),
            price,
            item.get("fohQuantity", 0),
            item.get("minOrderQuantity", ""),
            lead_time,
            item.get("warehouseCode", "")
        ])
        total_fetched += 1

wb.save("arrow_catalog.xlsx")
print(f"\nDone! {total_fetched} items saved to arrow_catalog.xlsx")
