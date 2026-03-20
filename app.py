from flask import Flask, jsonify, Response
from flask_cors import CORS
import requests
import openpyxl
import io
import os

app = Flask(__name__)
CORS(app)

# =============================================
# PUT YOUR CREDENTIALS HERE
MALABS_EMAIL = os.environ.get("MALABS_EMAIL", "your_email@example.com")
MALABS_PASSWORD = os.environ.get("MALABS_PASSWORD", "your_password")
# =============================================

@app.route("/")
def home():
    return jsonify({"status": "Distributor Portal Backend Running"})

@app.route("/malabs/status")
def malabs_status():
    try:
        url = "https://online.malabs.com/mws/items/?format=json&in_stock=Y&page=1"
        r = requests.get(url, auth=(MALABS_EMAIL, MALABS_PASSWORD), timeout=10)
        if r.status_code == 200:
            data = r.json()
            return jsonify({"status": "connected", "total_items": data.get("count", 0)})
        else:
            return jsonify({"status": "error", "code": r.status_code})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route("/malabs/fetch")
def malabs_fetch():
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MA Labs Catalog"
        ws.append(["Item No", "UPC Code", "Manufacturer No", "Manufacturer",
                   "Category", "Product Name", "Price", "Inventory"])

        page = 1
        total = 0

        while True:
            url = f"https://online.malabs.com/mws/items/?format=json&in_stock=Y&page={page}"
            r = requests.get(url, auth=(MALABS_EMAIL, MALABS_PASSWORD), timeout=30)
            if r.status_code != 200:
                break
            data = r.json()
            results = data.get("results", [])
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
            total += len(results)
            if not data.get("next"):
                break
            page += 1

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return Response(
            output.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment;filename=malabs_catalog.xlsx"}
        )

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
