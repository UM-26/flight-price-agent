import json
import os
from pathlib import Path
import requests

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

API = "https://serpapi.com/search.json"

params = {
    "engine": "google_flights",
    "api_key": os.environ["SERPAPI_KEY"],
    "departure_id": ",".join(CFG["departure_airports"]),
    "arrival_id": ",".join(CFG["arrival_airports"]),
    "outbound_date": "2027-05-12",
    "return_date": "2027-05-22",
    "type": "1",
    "adults": CFG["adults"],
    "children": CFG["children"],
    "infants_in_seat": 0,
    "travel_class": "1",
    "currency": CFG["currency"],
    "hl": CFG["language"],
    "gl": CFG["country"],
    "stops": CFG["stops"],
    "layover_duration": f'0,{CFG["max_layover_hours"] * 60}',
    "sort_by": "2",
}

print("Testing Google Flights...")
print("Route:", ",".join(CFG["departure_airports"]), "->", ",".join(CFG["arrival_airports"]))
print("Dates: 2027-05-12 -> 2027-05-22")
print("Passengers:", CFG["adults"], "adults +", CFG["children"], "child")
print("Filters: max 1 stop, max", CFG["max_layover_hours"], "hours layover")

response = requests.get(API, params=params, timeout=60)
response.raise_for_status()
data = response.json()

if data.get("error"):
    raise RuntimeError(data["error"])

best = data.get("best_flights") or []
other = data.get("other_flights") or []
all_flights = best + other

print("API OK")
print("Best flights:", len(best))
print("Other flights:", len(other))
print("Total itineraries:", len(all_flights))

if not all_flights:
    print("Response keys:", sorted(data.keys()))
    raise RuntimeError("Google Flights returned no matching itineraries.")

def price(item):
    value = item.get("price")
    return value if isinstance(value, (int, float)) else 10**12

all_flights.sort(key=price)

print("\nTop results:")
for i, item in enumerate(all_flights[:10], 1):
    print(
        f"{i}. {item.get('price')} {CFG['currency']} | "
        f"stops={item.get('number_of_stops', '?')} | "
        f"duration={item.get('total_duration', item.get('duration', '?'))}"
    )

    for direction in ("outbound", "return"):
        part = item.get(direction) or {}
        for segment in part.get("flights", [])[:5]:
            print(
                "   ",
                segment.get("departure_airport", {}).get("id", "?"),
                "->",
                segment.get("arrival_airport", {}).get("id", "?"),
                "|",
                segment.get("airline", "?"),
                "|",
                segment.get("departure_airport", {}).get("time", "?"),
                "->",
                segment.get("arrival_airport", {}).get("time", "?"),
            )

    layovers = item.get("layovers") or []
    if layovers:
        print("    layovers:", [
            {
                "airport": x.get("name"),
                "duration": x.get("duration"),
            }
            for x in layovers
        ])

print("\nTEST PASSED: Google Flights returned usable itineraries.")
