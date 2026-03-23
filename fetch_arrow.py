import requests
import os
from openpyxl import Workbook

# =============================================
# PASTE YOUR ACCESS TOKEN HERE (from Postman)
ACCESS_TOKEN = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzY29wZSI6WyJNeUFycm93RXh0ZXJuYWxDbGllbnQiXSwiZXhwIjoxNzc0MzEwNDc4LCJqdGkiOiJHbEZubG9uQy1mM0t4R2JJemlDeVpvRUl1YXciLCJjbGllbnRfaWQiOiJhbGdvcnl4LWxsYyJ9.PZFF24QOWzRsWUrTWvFgpHRZAdvautE3y37diWZhLnfzNbphO6pLxZU7pRgSGC5E78RPnguuRjmZWV5pAJl96YsQn7--zzpcngIGmifd-Z1NoWvn7IJQnOttXREQqdZn6bsElMG9S8zkuRKpnt072ODoFmujaiIrsdaQDQ7oU8SIJk2vP4oc9Wfk-LDw94LywIiwHd_RCFPKuRqMHaDWU7Ksk6OIb7nlj5A56QPCtfFIE2fVdOIg62GDwehwSJS-bPs17aJsQ71snC7PnXxwfKBwqHptYIyoRe006drBkkR7nEDWQ_r9zvzS1I3VGjFZt66JI31wM3tuJ5D6mlK5Zg"
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
