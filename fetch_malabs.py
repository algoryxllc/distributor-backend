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

# Store records for JSON sync
json_records = []

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

        # Collect for JSON sync
        json_records.append({
            "mfr_sku": item.get("manufacturer_no", ""),
            "disti_sku": item.get("item_no", ""),
            "manufacturer": item.get("manufacturer", ""),
            "price": item.get("price", ""),
            "quantity": sum([
                int(inventory.get("1001", 0) or 0),
                int(inventory.get("1002", 0) or 0),
                int(inventory.get("1003", 0) or 0),
                int(inventory.get("1004", 0) or 0),
                int(inventory.get("1005", 0) or 0),
                int(inventory.get("1006", 0) or 0),
            ]),
            "weight": item.get("weight", ""),
            "length": item.get("length", ""),
            "width": item.get("width", ""),
            "height": item.get("height", ""),
        })

        total_fetched += 1

    offset += 10
    print(f"Offset {offset}/{total_items} — {total_fetched} items fetched...")

    # Use next URL from API response
    next_url = page_data.get("next")
    if next_url and "format=json" not in next_url:
        next_url = next_url + "&format=json"

wb.save("malabs_catalog.xlsx")
print(f"\nDone! {total_fetched} items saved to malabs_catalog.xlsx")

# =============================================
# ADDED: Save JSON to GitHub repo for auto-sync
# Nothing above this line was changed
# =============================================
import json, base64, datetime

GH_TOKEN = os.environ.get("MALABS_PAT")
GH_OWNER = "algoryxllc"
GH_REPO  = "distributor-backend"
JSON_PATH = "data/malabs_latest.json"

if GH_TOKEN:
    print("\nSaving JSON to GitHub for Pricing Portal auto-sync...")

    output = {
        "fetched_at": datetime.datetime.utcnow().isoformat() + "Z",
        "distributor": "MA Labs",
        "total": total_fetched,
        "records": json_records
    }

    encoded = base64.b64encode(json.dumps(output).encode()).decode()

    # Check if file already exists (need SHA to update)
    check = requests.get(
        f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/contents/{JSON_PATH}",
        headers={"Authorization": f"token {GH_TOKEN}"}
    )
    sha = check.json().get("sha") if check.status_code == 200 else None

    payload = {
        "message": f"Auto-update MA Labs — {total_fetched} items",
        "content": encoded
    }
    if sha:
        payload["sha"] = sha

    result = requests.put(
        f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/contents/{JSON_PATH}",
        headers={"Authorization": f"token {GH_TOKEN}"},
        json=payload
    )

    if result.status_code in [200, 201]:
        print(f"JSON saved to GitHub: {JSON_PATH}")
    else:
        print(f"Failed to save JSON: {result.status_code}")
else:
    print("No GH_TOKEN found — skipping JSON save")
