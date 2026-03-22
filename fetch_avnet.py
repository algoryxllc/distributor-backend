import requests
import os
from openpyxl import Workbook
import time

CLIENT_ID = os.environ.get("AVNET_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AVNET_CLIENT_SECRET")
SKUS = os.environ.get("SKUS", "")

def get_token():
    print("Getting Avnet OAuth token...")
    response = requests.post(
        "https://apigw.avnet.com/oauth/client_credential/accesstoken",
        params={"grant_type": "client_credentials"},
        auth=(CLIENT_ID, CLIENT_SECRET),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=60
    )
    print(f"Token response: {response.status_code}")
    if response.status_code == 200:
        token = response.json().get("access_token")
        print("Token obtained successfully!")
        return token
    else:
        print(f"Token error: {response.status_code} - {response.text}")
        return None

def fetch_skus(token, sku_batch, start_id):
    url = "https://apigw.avnet.com/external/customer/price/v1/"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    items = []
    for i, sku in enumerate(sku_batch):
        items.append({
            "itemId": start_id + i,
            "searchType": "REQUEST_PART",
            "searchTerm": sku,
            "quantity": 1
        })

    payload = {
        "pageRows": 10,
        "pageNum": 1,
        "stock": "Y",
        "price": "Y",
        "items": items
    }

    response = requests.post(url, headers=headers, json=payload, timeout=60)
    if response.status_code == 200:
        return response.json()
    elif response.status_code == 429:
        print("Rate limit hit! Waiting 60 seconds...")
        time.sleep(60)
        return fetch_skus(token, sku_batch, start_id)
    else:
        print(f"Error: {response.status_code} - {response.text}")
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
ws.title = "Avnet Catalog"
ws.append([
    "Search SKU",
    "quotedPartNumber",
    "quotedManufacturerName",
    "materialDescription",
    "price",
    "currency",
    "inStock",
    "sellQuantity",
    "minimumQuantity",
    "multipleQuantity",
    "factoryLeadTimeWks",
    "obsoleteFlag",
    "endOfLife",
    "packageDescription",
    "countryOfOrigin",
    "rohsComplianceCode",
    "eccn",
    "expirationDate",
    "comments"
])

total_fetched = 0
batch_size = 10  # Max 10 per request as per document

for i in range(0, len(sku_list), batch_size):
    batch = sku_list[i:i + batch_size]
    print(f"Fetching batch {i//batch_size + 1}: SKUs {i+1}-{i+len(batch)}")

    data = fetch_skus(token, batch, i + 1)

    if not data:
        for sku in batch:
            ws.append([sku, "ERROR"] + [""] * 17)
        continue

    result_items = data.get("items", [])
    sku_map = {item["itemId"]: batch[item["itemId"] - i - 1] for item in result_items if item.get("itemId")}

    for item in result_items:
        item_id = item.get("itemId", 0)
        search_sku = batch[item_id - i - 1] if 0 <= (item_id - i - 1) < len(batch) else ""

        ws.append([
            search_sku,
            item.get("quotedPartNumber", ""),
            item.get("quotedManufacturerName", ""),
            item.get("materialDescription", ""),
            item.get("price", ""),
            item.get("currency", ""),
            item.get("inStock", ""),
            item.get("sellQuantity", ""),
            item.get("minimumQuantity", ""),
            item.get("multipleQuantity", ""),
            item.get("factoryLeadTimeWks", ""),
            item.get("obsoleteFlag", ""),
            item.get("endOfLife", ""),
            item.get("packageDescription", ""),
            item.get("countryOfOrigin", ""),
            item.get("rohsComplianceCode", ""),
            item.get("eccn", ""),
            item.get("expirationDate", ""),
            item.get("comments", "")
        ])
        total_fetched += 1

    # Respect rate limit - 500 calls per 5 min
    time.sleep(0.6)

wb.save("avnet_catalog.xlsx")
print(f"\nDone! {total_fetched} items saved to avnet_catalog.xlsx")
