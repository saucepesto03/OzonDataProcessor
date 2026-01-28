import requests
import pandas as pd
import os
import sys
from pathlib import Path

# ========== ИСПРАВЛЕНО ДЛЯ EXE ==========
if getattr(sys, 'frozen', False):
    # Запущено из .exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Запущено из .py
    BASE_DIR = Path(__file__).parent
# =======================================

BASE_URL = "https://api-seller.ozon.ru"

# ПУТИ ОТНОСИТЕЛЬНО РАБОЧЕЙ ДИРЕКТОРИИ
INPUT_API_FILE = BASE_DIR / "apis.txt"
OUTPUT_DIR = BASE_DIR / "ozon_dimensions"
OUTPUT_FILE = "ozon_dimensions_cm.xlsx"
OUTPUT_PATH = OUTPUT_DIR / OUTPUT_FILE

def mm_to_cm(value):
    return round(value / 10, 2) if isinstance(value, (int, float)) else None

def load_apis(filepath):
    """
    Формат apis.txt:
    client_id;api_key
    client_id;api_key
    """
    apis = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            client_id, api_key = line.split(";", 1)
            apis.append({
                "Client-Id": client_id.strip(),
                "Api-Key": api_key.strip(),
                "Content-Type": "Application/json"
            })
    return apis

def get_all_product_ids(headers):
    product_ids = []
    last_id = ""

    while True:
        payload = {
            "filter": {"visibility": "ALL"},
            "limit": 1000,
            "last_id": last_id
        }
        r = requests.post(f"{BASE_URL}/v3/product/list", json=payload, headers=headers)
        data = r.json()["result"]

        for item in data["items"]:
            product_ids.append(item["product_id"])

        last_id = data["last_id"]
        if not last_id:
            break

    return product_ids

def get_products_attributes(product_ids, headers):
    rows = []

    for i in range(0, len(product_ids), 100):
        chunk = product_ids[i:i + 100]
        payload = {
            "filter": {"product_id": chunk},
            "limit": 100
        }
        r = requests.post(
            f"{BASE_URL}/v4/product/info/attributes",
            json=payload,
            headers=headers
        )
        products = r.json()["result"]

        for p in products:
            rows.append({
                "sku": p.get("sku"),
                "name": p.get("name"),
                "offer_id": p.get("offer_id"),
                "width_cm": mm_to_cm(p.get("width")),
                "height_cm": mm_to_cm(p.get("height")),
                "length_cm": mm_to_cm(p.get("depth"))
            })

    return rows

def main():
    # Создаем папку если нет
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    if OUTPUT_PATH.exists():
        print("⚠️  отчёт уже существует")
        print(f"  Файл: {OUTPUT_PATH}")
        print("  Удалите его вручную или переименуйте, чтобы создать новый")
        return

    print(f"Рабочая директория: {BASE_DIR}")
    print(f"Файл API: {INPUT_API_FILE}")
    print(f"Выходная папка: {OUTPUT_DIR}")

    try:
        apis = load_apis(INPUT_API_FILE)
    except FileNotFoundError:
        print(f"❌ Файл {INPUT_API_FILE.name} не найден!")
        print(f"Создайте файл {INPUT_API_FILE.name} в директории:")
        print(f"  {BASE_DIR}")
        print("Формат файла:")
        print("  client_id;api_key")
        print("  client_id;api_key")
        return

    all_rows = []

    for idx, headers in enumerate(apis, start=1):
        print(f"\nОбработка API #{idx}")
        product_ids = get_all_product_ids(headers)
        print(f"  Найдено товаров: {len(product_ids)}")
        rows = get_products_attributes(product_ids, headers)
        all_rows.extend(rows)

    df = pd.DataFrame(all_rows)
    df.to_excel(OUTPUT_PATH, index=False)
    print(f"\n✅ Отчёт сохранен: {OUTPUT_PATH}")
    print(f"✅ Обработано товаров: {len(df)}")

if __name__ == "__main__":
    main()