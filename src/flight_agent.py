import csv
import json
import os
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
CSV = ROOT / "data" / "prices.csv"
XLSX = ROOT / "data" / "flight_prices.xlsx"
ALERT_STATE = ROOT / "data" / "alert_state.json"
ALERT_FILE = ROOT / "data" / "alert.md"
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


def build_combinations(all_rows):
    """
    Build historical best observed 6-8 night combinations without API calls.

    For each exact one-way itinerary/date, keep its lowest observed price.
    Then combine outbound and return observations whose dates are 6, 7 or 8
    nights apart. The result is a historical combination, not a live quote.
    """
    best = {}

    for row in all_rows:
        try:
            price = int(row.get("price_czk", ""))
            flight_date = row.get("flight_date", "")
        except (TypeError, ValueError):
            continue

        if row.get("direction") not in {"OUTBOUND", "RETURN"} or not flight_date:
            continue

        key = (
            row["direction"],
            flight_date,
            row.get("origin", ""),
            row.get("destination", ""),
            row.get("airline", ""),
            row.get("flight_numbers", ""),
        )

        if key not in best or price < best[key]["price_czk"]:
            item = dict(row)
            item["price_czk"] = price
            best[key] = item

    outbound = [item for item in best.values() if item["direction"] == "OUTBOUND"]
    returns = [item for item in best.values() if item["direction"] == "RETURN"]

    by_return_date = {}
    for row in returns:
        by_return_date.setdefault(row["flight_date"], []).append(row)

    combinations = []

    for out in outbound:
        try:
            out_date = date.fromisoformat(out["flight_date"])
        except ValueError:
            continue

        for nights in range(CFG["trip_length_min"], CFG["trip_length_max"] + 1):
            return_date = out_date + timedelta(days=nights)
            return_rows = by_return_date.get(return_date.isoformat(), [])

            for ret in return_rows:
                total = int(out["price_czk"]) + int(ret["price_czk"])
                same_nyc_airport = out["destination"] == ret["origin"]

                combinations.append({
                    "odlet": out["flight_date"],
                    "návrat": ret["flight_date"],
                    "noci": nights,
                    "celkem_czk": total,
                    "odlet_z": out["origin"],
                    "prilet_do_nyc": out["destination"],
                    "navrat_z_nyc": ret["origin"],
                    "prilet_do": ret["destination"],
                    "cena_odlet_czk": int(out["price_czk"]),
                    "cena_navrat_czk": int(ret["price_czk"]),
                    "letec_odlet": out["airline"],
                    "letec_navrat": ret["airline"],
                    "prestupy_odlet": int(out.get("stops") or 0),
                    "prestupy_navrat": int(ret.get("stops") or 0),
                    "max_prestup_odlet_min": int(out.get("max_layover_minutes") or 0),
                    "max_prestup_navrat_min": int(ret.get("max_layover_minutes") or 0),
                    "poznamka": "" if same_nyc_airport else "Jiná NYC letiště pro přílet a odlet",
                    "poznamka_typ": "Historicky nejnižší zaznamenané ceny pro dané jednosměrné itineráře",
                })

    combinations.sort(key=lambda row: row["celkem_czk"])
    return combinations[:2000]


def combo_key(row):
    return (
        row["odlet"],
        row["návrat"],
        row["odlet_z"],
        row["prilet_do_nyc"],
        row["navrat_z_nyc"],
        row["prilet_do"],
        row["noci"],
    )


def load_alert_state():
    if not ALERT_STATE.exists():
        return {}
    try:
        return json.loads(ALERT_STATE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def build_alert(combinations):
    """
    Compare the current historical best with the previous best.

    First run creates a baseline and does not alert. Later runs alert when:
    - the overall best combination gets cheaper by at least the configured
      absolute or percentage drop, or
    - the best price crosses the configured absolute alert threshold.
    """
    if not combinations:
        return None

    best = combinations[0]
    state = load_alert_state()
    previous_price = state.get("best_price_czk")

    current_price = int(best["celkem_czk"])
    drop_czk = 0
    drop_percent = 0.0
    alert_reason = []

    if isinstance(previous_price, (int, float)) and current_price < previous_price:
        drop_czk = int(previous_price - current_price)
        drop_percent = (drop_czk / previous_price) * 100 if previous_price else 0
        if drop_czk >= CFG.get("alert_drop_czk", 1000):
            alert_reason.append(f"pokles o {drop_czk:,} Kč".replace(",", " "))
        if drop_percent >= CFG.get("alert_drop_percent", 5):
            alert_reason.append(f"pokles o {drop_percent:.1f} %")

    threshold = CFG.get("alert_price_czk")
    if isinstance(threshold, (int, float)) and current_price <= threshold:
        if not state.get("threshold_alerted", False):
            alert_reason.append(f"cena je pod hranicí {int(threshold):,} Kč".replace(",", " "))

    state_update = {
        "best_price_czk": current_price,
        "best_key": list(combo_key(best)),
        "last_checked_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "threshold_alerted": bool(
            isinstance(threshold, (int, float)) and current_price <= threshold
        ),
    }

    ALERT_STATE.parent.mkdir(parents=True, exist_ok=True)
    ALERT_STATE.write_text(
        json.dumps(state_update, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    if not alert_reason:
        if not previous_price:
            print(f"Alert baseline created: {current_price} CZK")
        return None

    reason = " a ".join(dict.fromkeys(alert_reason))
    previous_text = (
        f"{previous_price:,} Kč".replace(",", " ")
        if isinstance(previous_price, (int, float))
        else "není k dispozici"
    )

    body = f"""# 🚨 Nová zajímavá cena do New Yorku

**Důvod:** {reason}

## Nejlepší nalezená kombinace

- **Celkem:** {current_price:,} Kč
- **Odlet:** {best["odlet"]} z {best["odlet_z"]} → {best["prilet_do_nyc"]}
- **Návrat:** {best["návrat"]} z {best["navrat_z_nyc"]} → {best["prilet_do"]}
- **Pobyt:** {best["noci"]} nocí
- **Odlet:** {best["cena_odlet_czk"]:,} Kč, {best["letec_odlet"]}, {best["prestupy_odlet"]} přestup(y)
- **Návrat:** {best["cena_navrat_czk"]:,} Kč, {best["letec_navrat"]}, {best["prestupy_navrat"]} přestup(y)
- **Předchozí nejlepší cena:** {previous_text}

> ⚠️ Toto je historicky zaznamenaná kombinace. Před rezervací je nutné ověřit aktuální dostupnost a cenu v Google Flights.

**Kontrola:** {state_update["last_checked_at_utc"]}
"""
    ALERT_FILE.write_text(body, encoding="utf-8")
    return body


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

    combinations = build_combinations(all_rows)

    wb = Workbook()
    dashboard = wb.active
    dashboard.title = "Přehled"

    dashboard["A1"] = "✈️ HLÍDAČ NEW YORKU"
    dashboard["A1"].font = Font(size=18, bold=True)
    dashboard.merge_cells("A1:H1")

    dashboard["A3"] = "Nejlepší historicky zaznamenaná kombinace"
    dashboard["A3"].font = Font(size=13, bold=True)

    if combinations:
        best = combinations[0]
        dashboard["A4"] = "Celkem"
        dashboard["B4"] = best["celkem_czk"]
        dashboard["B4"].font = Font(size=16, bold=True)
        dashboard["A5"] = "Termín"
        dashboard["B5"] = f'{best["odlet"]} → {best["návrat"]} ({best["noci"]} nocí)'
        dashboard["A6"] = "Trasa"
        dashboard["B6"] = f'{best["odlet_z"]} → {best["prilet_do_nyc"]} / {best["navrat_z_nyc"]} → {best["prilet_do"]}'
        dashboard["A7"] = "Let tam"
        dashboard["B7"] = f'{best["cena_odlet_czk"]} Kč | {best["letec_odlet"]} | {best["prestupy_odlet"]} přestup(y)'
        dashboard["A8"] = "Let zpět"
        dashboard["B8"] = f'{best["cena_navrat_czk"]} Kč | {best["letec_navrat"]} | {best["prestupy_navrat"]} přestup(y)'
        dashboard["A10"] = "Důležité"
        dashboard["B10"] = "Cena je historicky pozorovaná kombinace, nikoli živá nabídka."
    else:
        dashboard["A4"] = "Zatím není dost dat pro kombinaci 6–8 nocí."

    dashboard["A12"] = "Co hlídám"
    dashboard["A12"].font = Font(size=13, bold=True)
    dashboard["A13"] = "Cestující"
    dashboard["B13"] = f'{CFG["adults"]} dospělí + {CFG["children"]} dítě'
    dashboard["A14"] = "Odlety"
    dashboard["B14"] = ", ".join(ORIGINS)
    dashboard["A15"] = "New York"
    dashboard["B15"] = ", ".join(NYC)
    dashboard["A16"] = "Pobyt"
    dashboard["B16"] = f'{CFG["trip_length_min"]}–{CFG["trip_length_max"]} nocí'
    dashboard["A17"] = "Přestupy"
    dashboard["B17"] = f'0–1, max. {CFG["max_layover_hours"]} h'
    dashboard["A18"] = "Období"
    dashboard["B18"] = f'{CFG["outbound_start"]} → {CFG["outbound_end"]}'

    dashboard.column_dimensions["A"].width = 38
    dashboard.column_dimensions["B"].width = 72
    dashboard.freeze_panes = "A3"

    ws = wb.create_sheet("Historie letů")
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

    summary = wb.create_sheet("Souhrn")
    summary.append(["Položka", "Hodnota"])
    summary.append(["Počet uložených letů", len(all_rows)])

    if all_rows:
        prices = [int(row["price_czk"]) for row in all_rows if row.get("price_czk")]
        summary.append(["Nejnižší zaznamenaná cena jednosměrně (Kč)", min(prices)])

    summary.append(["Nejlepší historická kombinace (Kč)", combinations[0]["celkem_czk"] if combinations else ""])
    summary.append(["Počet kombinací 6–8 nocí", len(combinations)])
    summary.append(["Poslední kontrola (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")])
    summary.append(["Cíl cesty", "New York"])
    summary.append(["Délka pobytu", "6–8 nocí"])
    summary.append(["Odletové období", f'{CFG["outbound_start"]} až {CFG["outbound_end"]}'])
    summary.append(["Přestupy", "0 nebo 1"])
    summary.append(["Max. délka přestupu", f'{CFG["max_layover_hours"]} hodin'])
    summary.column_dimensions["A"].width = 42
    summary.column_dimensions["B"].width = 34

    combo_ws = wb.create_sheet("Kombinace 6-8 nocí")
    combo_fields = [
        "odlet", "návrat", "noci", "celkem_czk", "odlet_z", "prilet_do_nyc",
        "navrat_z_nyc", "prilet_do", "cena_odlet_czk", "cena_navrat_czk",
        "letec_odlet", "letec_navrat", "prestupy_odlet", "prestupy_navrat",
        "max_prestup_odlet_min", "max_prestup_navrat_min", "poznamka", "poznamka_typ",
    ]
    combo_ws.append(combo_fields)
    for row in combinations:
        combo_ws.append([row.get(field, "") for field in combo_fields])
    combo_ws.freeze_panes = "A2"
    combo_ws.auto_filter.ref = combo_ws.dimensions

    combo_widths = {
        "A": 14, "B": 14, "C": 8, "D": 14, "E": 10, "F": 15,
        "G": 15, "H": 10, "I": 16, "J": 17, "K": 22, "L": 22,
        "M": 16, "N": 17, "O": 23, "P": 24, "Q": 34, "R": 55,
    }
    for column, width in combo_widths.items():
        combo_ws.column_dimensions[column].width = width

    combo_summary = wb.create_sheet("Jak číst kombinace")
    combo_summary.append(["Informace", "Vysvětlení"])
    combo_summary.append([
        "Co obsahuje list",
        "Kombinace historicky zaznamenaných jednosměrných letů pro 6, 7 nebo 8 nocí.",
    ])
    combo_summary.append([
        "Cena",
        "Celková cena = zaznamenaná cena odletu + zaznamenaná cena návratu pro 2 dospělé a 1 dítě.",
    ])
    combo_summary.append([
        "Důležité",
        "Jde o historicky pozorované ceny, nikoli o aktuální nabídku dostupnou k okamžité rezervaci.",
    ])
    combo_summary.append([
        "NYC letiště",
        "Pokud je přílet do jiného NYC letiště než odlet zpět, je to označeno v poznámce.",
    ])
    combo_summary.append([
        "Řazení",
        "Nejnižší historicky zaznamenané celkové kombinace jsou nahoře.",
    ])
    combo_summary.column_dimensions["A"].width = 22
    combo_summary.column_dimensions["B"].width = 100

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

    all_rows = []
    with CSV.open("r", encoding="utf-8", newline="") as f:
        all_rows = list(csv.DictReader(f))

    combinations = build_combinations(all_rows)
    alert = build_alert(combinations)

    print(f"Saved {len(rows)} flight results.")
    print(
        f'Best one-way: {rows[0]["price_czk"]} CZK | '
        f'{rows[0]["origin"]} -> {rows[0]["destination"]} | '
        f'{rows[0]["flight_date"]}'
    )

    if combinations:
        print(
            f'Best historical 6-8 night combination: '
            f'{combinations[0]["celkem_czk"]} CZK | '
            f'{combinations[0]["odlet"]} -> {combinations[0]["návrat"]}'
        )

    if alert:
        print("PRICE ALERT: new interesting combination detected.")
    else:
        print("No new price alert.")


if __name__ == "__main__":
    main()
