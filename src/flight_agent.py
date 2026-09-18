import csv
import json
import os
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
CSV = ROOT / "data" / "prices.csv"
XLSX = ROOT / "data" / "flight_prices.xlsx"
API = "https://serpapi.com/search.json"

ORIGINS = CFG["departure_airports"]
NYC = CFG["arrival_airports"]


def daterange(start, end):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def choose_dates(today):
    """Pick one outbound date and one 6-8 night return date for this run."""
    start = date.fromisoformat(CFG["outbound_start"])
    end = date.fromisoformat(CFG["outbound_end"])
    dates = list(daterange(start, end))

    if not dates:
        raise RuntimeError("No outbound dates configured.")

    index = (today - start).days % len(dates)
    outbound = dates[index]

    stay_span = CFG["trip_length_max"] - CFG["trip_length_min"] + 1
    nights = CFG["trip_length_min"] + ((today - start).days % stay_span)
    return outbound, outbound + timedelta(days=nights)


def request_flights(key, departure_id, arrival_id, outbound_date):
    params = {
        "engine": "google_flights",
        "api_key": key,
        "departure_id": departure_id,
        "arrival_id": arrival_id,
        "outbound_date": outbound_date.isoformat(),
        "type": "2",
        "adults": CFG["adults"],
        "children": CFG["children"],
        "infants_on_lap": CFG.get("infants_on_lap", 0),
        "travel_class": "1",
        "currency": CFG["currency"],
        "hl": CFG["language"],
        "gl": CFG["country"],
        "stops": CFG["stops"],
        "layover_duration": f'0,{CFG["max_layover_hours"] * 60}',
        "sort_by": "2",
    }

    response = requests.get(API, params=params, timeout=90)
    response.raise_for_status()
    data = response.json()

    if data.get("error"):
        raise RuntimeError(data["error"])

    return data


def extract_itineraries(data):
    results = []
    for key in ("best_flights", "other_flights"):
        items = data.get(key, [])
        if isinstance(items, list):
            results.extend(item for item in items if isinstance(item, dict))
    return results


def leg_info(item):
    flights = item.get("flights", [])
    if not isinstance(flights, list):
        flights = []

    stops = max(0, len(flights) - 1)
    layovers = item.get("layovers", [])
    if not isinstance(layovers, list):
        layovers = []

    durations = [
        int(x.get("duration"))
        for x in layovers
        if isinstance(x, dict) and isinstance(x.get("duration"), (int, float))
    ]

    max_layover = max(durations) if durations else 0
    layover_airports = [
        x.get("id") or x.get("name", "")
        for x in layovers
        if isinstance(x, dict)
    ]

    if flights:
        first = flights[0]
        last = flights[-1]
        origin = first.get("departure_airport", {}).get("id", "")
        destination = last.get("arrival_airport", {}).get("id", "")
        airline = first.get("airline", "")
    else:
        origin = destination = airline = ""

    flight_numbers = [
        x.get("flight_number", "")
        for x in flights
        if isinstance(x, dict) and x.get("flight_number")
    ]

    return {
        "origin": origin,
        "destination": destination,
        "airline": airline,
        "stops": stops,
        "max_layover_minutes": max_layover,
        "layover_airports": ", ".join(layover_airports),
        "duration_minutes": item.get("total_duration", ""),
        "flight_numbers": ", ".join(flight_numbers),
    }


def collect_one_way(data, checked_at, direction, searched_date):
    rows = []

    for item in extract_itineraries(data):
        price = item.get("price")
        if not isinstance(price, (int, float)):
            continue

        info = leg_info(item)
        if not info["origin"] or not info["destination"]:
            continue

        if info["stops"] > 1 or info["max_layover_minutes"] > CFG["max_layover_hours"] * 60:
            continue

        rows.append({
            "checked_at_utc": checked_at,
            "direction": direction,
            "price_czk": int(price),
            "origin": info["origin"],
            "destination": info["destination"],
            "airline": info["airline"],
            "flight_date": searched_date.isoformat(),
            "stops": info["stops"],
            "max_layover_minutes": info["max_layover_minutes"],
            "layover_airports": info["layover_airports"],
            "duration_minutes": info["duration_minutes"],
            "flight_numbers": info["flight_numbers"],
            "source": "Google Flights via SerpApi",
        })

    return rows


def append_rows(rows):
    fields = [
        "checked_at_utc",
        "direction",
        "price_czk",
        "origin",
        "destination",
        "airline",
        "flight_date",
        "stops",
        "max_layover_minutes",
        "layover_airports",
        "duration_minutes",
        "flight_numbers",
        "source",
    ]

    CSV.parent.mkdir(parents=True, exist_ok=True)
    existing_header = None

    if CSV.exists() and CSV.stat().st_size > 0:
        with CSV.open("r", encoding="utf-8", newline="") as f:
            existing_header = next(csv.reader(f), None)

    if existing_header != fields:
        old_rows = []
        if CSV.exists() and CSV.stat().st_size > 0:
            with CSV.open("r", encoding="utf-8", newline="") as f:
                old_rows = list(csv.DictReader(f))

        with CSV.open("w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fields)
            writer.writeheader()
            writer.writerows(
                {field: row.get(field, "") for field in fields}
                for row in old_rows
                if "direction" in row
            )

    with CSV.open("a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writerows(rows)


def build_xlsx():
    fields = [
        "checked_at_utc",
        "direction",
        "price_czk",
        "origin",
        "destination",
        "airline",
        "flight_date",
        "stops",
        "max_layover_minutes",
        "layover_airports",
        "duration_minutes",
        "flight_numbers",
        "source",
    ]

    all_rows = []
    if CSV.exists():
        with CSV.open("r", encoding="utf-8", newline="") as f:
            all_rows = list(csv.DictReader(f))

    wb = Workbook()
    ws = wb.active
    ws.title = "Historie letů"
    ws.append(fields)

    for row in all_rows:
        ws.append([row.get(field, "") for field in fields])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 22, "B": 12, "C": 12, "D": 10, "E": 12, "F": 22,
        "G": 14, "H": 8, "I": 20, "J": 28, "K": 18, "L": 32, "M": 30,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    today_rows = [
        row for row in all_rows
        if row.get("checked_at_utc", "").startswith(date.today().isoformat())
    ]

    summary = wb.create_sheet("Souhrn")
    summary.append(["Položka", "Hodnota"])
    summary.append(["Počet uložených letů", len(all_rows)])

    if all_rows:
        prices = [int(row["price_czk"]) for row in all_rows if row.get("price_czk")]
        summary.append(["Nejnižší zaznamenaná cena (Kč)", min(prices)])

    if today_rows:
        today_prices = [int(row["price_czk"]) for row in today_rows]
        summary.append(["Nejnižší cena dnešní kontroly (Kč)", min(today_prices)])
    else:
        summary.append(["Nejnižší cena dnešní kontroly (Kč)", ""])

    summary.append(["Poslední kontrola (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")])
    summary.append(["Cíl cesty", "New York"])
    summary.append(["Délka pobytu", "6-8 nocí"])
    summary.append(["Odletové období", f'{CFG["outbound_start"]} až {CFG["outbound_end"]}'])
    summary.append(["Přestupy", "0 nebo 1"])
    summary.append(["Max. délka přestupu", f'{CFG["max_layover_hours"]} hodin'])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 32

    wb.save(XLSX)


def main():
    key = os.environ.get("SERPAPI_KEY")
    if not key:
        raise RuntimeError("SERPAPI_KEY is not set")

    today = date.today()
    outbound_date, return_date = choose_dates(today)
    checked_at = datetime.now(timezone.utc).isoformat(timespec="seconds")

    print(f"Rotation outbound date: {outbound_date}")
    print(f"Rotation stay length: {(return_date - outbound_date).days} nights")
    print(f"Return-side date: {return_date}")
    print("Searching outbound: all departure airports -> all NYC airports")
    print("Searching return: all NYC airports -> all departure airports")

    outbound_data = request_flights(
        key,
        ",".join(ORIGINS),
        ",".join(NYC),
        outbound_date,
    )
    return_data = request_flights(
        key,
        ",".join(NYC),
        ",".join(ORIGINS),
        return_date,
    )

    rows = []
    rows.extend(collect_one_way(outbound_data, checked_at, "OUTBOUND", outbound_date))
    rows.extend(collect_one_way(return_data, checked_at, "RETURN", return_date))

    if not rows:
        raise RuntimeError("Google Flights returned no usable itineraries for this rotation.")

    rows.sort(key=lambda row: row["price_czk"])
    rows = rows[: CFG.get("max_results", 50)]

    append_rows(rows)
    build_xlsx()

    print(f"Saved {len(rows)} flight results.")
    print(
        f'Best one-way: {rows[0]["price_czk"]} CZK | '
        f'{rows[0]["origin"]} -> {rows[0]["destination"]} | '
        f'{rows[0]["flight_date"]}'
    )


if __name__ == "__main__":
    main()
