import requests
import os
from openpyxl import Workbook

CLIENT_ID = os.environ.get("ARROW_CLIENT_ID")
CLIENT_SECRET = os.environ.get("ARROW_CLIENT_SECRET")
SKUS = os.environ.get("SKUS", "")

def get_token():
    print("Getting Arrow API token...")
    # grant_type as query param, Basic Auth with client_id header
    response = requests.post(
        "https://my.arrow.com/api/security/oauth/token",
        params={"grant_type": "client_credentials"},
        headers={"client_id": CLIENT_ID},
        auth=(CLIENT_ID, CLIENT_SECRET),
        timeout=30
    )
    print(f"Token response status: {response.status_code}")
    if response.status_code == 200:
        token = response.json().get("access_token")
        print("Token obtained successfully!")
        return token
    else:
        print(f"Token error: {response.status_code} - {response.text}")
        return None

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
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching {sku}: {response.status_code} - {response.text}")
        return None

# Parse SKUs
sku_list = [s.strip() for s in SKUS.strip().split("\n") if s.strip()]
print(f"Total SKUs to fetch: {len(sku_list)}")

# Get token
token = get_token()
if not token:
    print("Failed to get token. Exiting.")
    exit(1)

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
    data = fetch_part(token, sku)

    if not data:
        ws.append([sku, "ERROR", "", "", "", "", "", "", ""])
        continue

    pricing = data.get("pricingResponse", [])

    if not pricing:
        ws.append([sku, "NOT FOUND", "", "", "", "", "", "", ""])
        continue

    for item in pricing:
        part_number = item.get("partNumber", "")
        manufacturer = item.get("manufacturer", "")
        description = item.get("description", "")
        inventory = item.get("fohQuantity", 0)
        min_qty = item.get("minOrderQuantity", "")
        warehouse = item.get("warehouseCode", "")

        lead_time_obj = item.get("leadTime", {})
        lead_time = lead_time_obj.get("arrowLeadTime", "") if lead_time_obj else ""

        pricing_tiers = item.get("pricingTier", [])
        price = pricing_tiers[0].get("resalePrice", "") if pricing_tiers else ""

        ws.append([
            sku, part_number, manufacturer, description,
            price, inventory, min_qty, lead_time, warehouse
        ])
        total_fetched += 1

wb.save("arrow_catalog.xlsx")
print(f"\nDone! Total {total_fetched} items saved to arrow_catalog.xlsx")
