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


def parse_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def main():
    key = os.environ.get("SERPAPI_KEY")
    if not key:
        raise RuntimeError("SERPAPI_KEY is not set")

    # Deals API supports flexible outbound dates and a natural-language
    # destination query. We filter the returned deals to JFK/EWR/LGA below.
    params = {
        "engine": "google_flights_deals",
        "api_key": key,
        "departure_id": ",".join(CFG["departure_airports"]),
        "query": "New York",
        "outbound_date": f'{CFG["outbound_start"]},{CFG["outbound_end"]}',
        "type": "1",
        "adults": CFG["adults"],
        "children": CFG["children"],
        "infants_on_lap": CFG.get("infants_on_lap", 0),
        "travel_class": "1",
        "currency": CFG["currency"],
        "hl": CFG["language"],
        "gl": CFG["country"],
        "stops": CFG["stops"],
    }

    r = requests.get(API, params=params, timeout=60)
    r.raise_for_status()
    data = r.json()

    if data.get("error"):
        raise RuntimeError(data["error"])

    deals = data.get("deals")
    if not isinstance(deals, list):
        deals = []

    target_airports = set(CFG["arrival_airports"])
    min_days = CFG["trip_length_min"]
    max_days = CFG["trip_length_max"]

    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    rows = []

    for deal in deals:
        if not isinstance(deal, dict):
            continue

        price = deal.get("price")
        if not isinstance(price, (int, float)):
            continue

        origin = deal.get("departure_airport_code", "")
        destination = deal.get("arrival_airport_code", "")

        if origin not in CFG["departure_airports"]:
            continue
        if destination not in target_airports:
            continue

        outbound_date = deal.get("start_date", "")
        return_date = deal.get("end_date", "")
        out_day = parse_date(outbound_date)
        ret_day = parse_date(return_date)

        if not out_day or not ret_day:
            continue

        trip_days = (ret_day - out_day).days
        if not (min_days <= trip_days <= max_days):
            continue

        stops = deal.get("stops", "")
        try:
            stops_value = int(stops)
        except (ValueError, TypeError):
            stops_value = ""

        if isinstance(stops_value, int) and stops_value > 1:
            continue

        rows.append({
            "checked_at_utc": now,
            "price_czk": int(price),
            "origin": origin,
            "destination": destination,
            "airline": deal.get("airline", ""),
            "outbound_date": outbound_date,
            "return_date": return_date,
            "stops": stops_value,
            "max_layover_minutes": "",
            "max_layover": "",
            "duration_minutes": deal.get("flight_duration", ""),
            "source": "Google Flights via SerpApi",
        })

    if not rows:
        print("API call succeeded, but no matching NYC itinerary was returned.")
        print("API response keys:", sorted(data.keys()))
        print("Number of deals returned:", len(deals))
        if deals:
            sample = deals[0]
            print("First deal sample:", {
                k: sample.get(k)
                for k in (
                    "name", "price", "start_date", "end_date",
                    "departure_airport_code", "arrival_airport_code",
                    "stops", "airline",
                )
            })
        return

    rows.sort(key=lambda x: x["price_czk"])
    rows = rows[:CFG.get("max_results", 20)]

    CSV.parent.mkdir(parents=True, exist_ok=True)
    fields = list(rows[0].keys())
    exists = CSV.exists()

    with CSV.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        if not exists:
            writer.writeheader()
        writer.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Historie cen"
    ws.append(fields)

    with CSV.open("r", encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            ws.append([row.get(field, "") for field in fields])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 24, "B": 14, "C": 10, "D": 12, "E": 22,
        "F": 14, "G": 14, "H": 10, "I": 20, "J": 16,
        "K": 18, "L": 30,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    summary = wb.create_sheet("Souhrn")
    summary.append(["Položka", "Hodnota"])
    summary.append(["Počet uložených záznamů", ws.max_row - 1])
    summary.append([
        "Nejnižší zaznamenaná cena (Kč)",
        min(int(row["price_czk"]) for row in rows),
    ])
    summary.append(["Nejnižší cena dnešní kontroly (Kč)", rows[0]["price_czk"]])
    summary.append(["Poslední kontrola (UTC)", now])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 28

    wb.save(XLSX)

    print(
        f'Best price: {rows[0]["price_czk"]} CZK, '
        f'{rows[0]["origin"]} -> {rows[0]["destination"]}, '
        f'{rows[0]["outbound_date"]} to {rows[0]["return_date"]}'
    )


if __name__ == "__main__":
    main()
