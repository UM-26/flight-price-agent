# Flight Price Agent

Daily tracker for a New York trip.

Current search:
- Departures: PRG (Prague), VIE (Vienna), KTW (Katowice), BTS (Bratislava), KRK (Krakow)
- New York airports: JFK, EWR, LGA
- Departure window: 1 April – 30 June 2027
- Trip length: 7–14 days
- 2 adults + 1 child
- Economy, CZK
- Direct flights only for now

The date window is supported by SerpApi's Google Flights Deals API, which accepts a flexible outbound-date range together with a custom trip-length range. citeturn0search0

Data source: Google Flights through SerpApi.

Required GitHub Actions secret: SERPAPI_KEY.

Outputs:
- data/prices.csv
- data/flight_prices.xlsx
