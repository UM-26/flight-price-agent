import csv, json, os
from datetime import datetime, timezone
from pathlib import Path
import requests
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
CSV = ROOT / "data" / "prices.csv"
XLSX = ROOT / "data" / "flight_prices.xlsx"
API = "https://serpapi.com/search.json"

def main():
    key = os.environ.get("SERPAPI_KEY")
    if not key:
        raise RuntimeError("SERPAPI_KEY is not set")
    params = {
        "engine": "google_flights_deals",
        "api_key": key,
        "departure_id": ",".join(CFG["departure_airports"]),
        "arrival_id": ",".join(CFG["arrival_airports"]),
        "outbound_date": f'{CFG["outbound_start"]},{CFG["outbound_end"]}',
        "trip_length": f'{CFG["trip_length_min"]},{CFG["trip_length_max"]}',
        "type": "1",
        "adults": CFG["adults"], "children": CFG["children"],
        "travel_class": "1", "currency": CFG["currency"],
        "hl": CFG["language"], "gl": CFG["country"], "stops": CFG["stops"]
    }
    r = requests.get(API, params=params, timeout=60)
    r.raise_for_status()
    data = r.json()
    if data.get("error"):
        raise RuntimeError(data["error"])

    items = []
    for key_name in ("best_flights", "other_flights", "flights", "deals"):
        if isinstance(data.get(key_name), list):
            items.extend(data[key_name])

    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    rows = []
    for item in items:
        price = item.get("price")
        if not isinstance(price, (int, float)):
            continue
        out = item.get("outbound") or {}
        ret = item.get("return") or {}
        rows.append({
            "checked_at_utc": now,
            "price_czk": int(price),
            "origin": item.get("departure_airport", {}).get("id") or out.get("departure_airport", {}).get("id", ""),
            "destination": item.get("arrival_airport", {}).get("id") or out.get("arrival_airport", {}).get("id", ""),
            "airline": item.get("airline", ""),
            "outbound_date": item.get("outbound_date") or out.get("date", ""),
            "return_date": item.get("return_date") or ret.get("date", ""),
            "stops": item.get("number_of_stops", item.get("stops", "")),
            "duration_minutes": item.get("duration", item.get("flight_duration", "")),
            "source": "Google Flights via SerpApi"
        })

    if not rows:
        print("No priced results. API keys:", sorted(data.keys()))
        return

    rows.sort(key=lambda x: x["price_czk"])
    rows = rows[:CFG.get("max_results", 20)]
    CSV.parent.mkdir(parents=True, exist_ok=True)
    fields = list(rows[0].keys())
    exists = CSV.exists()
    with CSV.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        if not exists: w.writeheader()
        w.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Historie cen"
    ws.append(fields)
    with CSV.open("r", encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            ws.append([row.get(x, "") for x in fields])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, width in enumerate([24,14,10,12,22,14,14,10,18,28], 1):
        ws.column_dimensions[chr(64+i)].width = width

    summary = wb.create_sheet("Souhrn")
    summary.append(["Položka", "Hodnota"])
    summary.append(["Počet uložených záznamů", ws.max_row - 1])
    summary.append(["Nejnižší zaznamenaná cena (Kč)", min(x["price_czk"] for x in rows)])
    summary.append(["Nejnižší cena dnešní kontroly (Kč)", rows[0]["price_czk"]])
    summary.append(["Poslední kontrola (UTC)", now])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 28
    wb.save(XLSX)
    print(f'Best price: {rows[0]["price_czk"]} CZK, {rows[0]["origin"]} -> {rows[0]["destination"]}')

if __name__ == "__main__":
    main()
