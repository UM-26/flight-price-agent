# Flight Price Agent

Automated price history for a surprise New York trip.

## Current search rules

- Departure airports: PRG, VIE, KTW, BTS, KRK
- New York airports: JFK, EWR, LGA
- Outbound dates: **24 April – 30 June 2027**
- Stay target: **6–8 nights**
- Passengers: 2 adults + 1 child
- Economy
- CZK
- 0 or 1 stop
- Maximum layover: 5 hours
- One-way searches are used separately for outbound and return flights
- The search rotates through the outbound date window instead of checking every date on every run

## How the rotation works

Each scheduled run makes two Google Flights searches:

1. All departure airports -> all NYC airports for one outbound date.
2. All NYC airports -> all departure airports for the corresponding return-side date.

The selected outbound date rotates through 24 April – 30 June 2027. The return-side search is initially 7 nights later. As the rotation progresses, the saved history can be combined to evaluate 6–8 night stays.

This keeps the number of API calls low while building a useful historical price dataset.

SerpApi's Google Flights API supports one-way searches with `type=2`, multiple airport IDs, passenger counts, stop filters and layover-duration filters. Cached identical searches are not counted toward the monthly search limit. citeturn0search0

## Outputs

- `data/prices.csv` - raw historical flight observations
- `data/flight_prices.xlsx` - Excel workbook with history and summary

## GitHub Actions

The workflow runs three times per week and can also be started manually from GitHub Actions.

Required repository secret:

`SERPAPI_KEY`

## Local test

```powershell
python src/test_google_flights.py
```

Data source: Google Flights through SerpApi.
