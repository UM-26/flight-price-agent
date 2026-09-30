# Flight Price Agent

Automated price history and price alerts for a surprise New York trip.

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

## What the agent does

Each scheduled run makes three API queries:

1. Smoke test for one-way Google Flights.
2. Outbound search for one selected date.
3. Return search for the corresponding 6–8 night rotation date.

The production searches are still only two Google Flights queries. The smoke test is a separate validation query.

The selected outbound date rotates through 24 April – 30 June 2027. The return-side search uses a 6, 7 or 8 night stay. Over time, the saved one-way history can therefore be combined into complete 6–8 night trip options without extra API calls.

## The useful part: trip combinations

The workbook now starts with **Přehled**, which shows the lowest historically observed complete 6–8 night combination found so far.

It includes:

- total price for 2 adults + 1 child
- outbound and return dates
- departure and NYC airports
- outbound and return airlines
- number of stops
- search constraints

The **Kombinace 6-8 nocí** sheet contains the complete historical ranking.

Important: these are historical observations assembled from one-way searches. They are not guaranteed to be simultaneously bookable at the displayed total price. Before booking, the current Google Flights result must be checked.

## Price alerts

The agent keeps the file data/alert_state.json with the previous best historical combination.

A GitHub alert is created when:

- the best total price falls by at least **1,000 Kč**, or
- it falls by at least **5 %**, or
- the price crosses the configured absolute threshold in config.json.

The first run only creates the baseline and does not send an alert.

Alerts are created as GitHub Issues, so GitHub can notify you according to your repository notification/watch settings.

### Important price detail

The Google Flights query is made for **2 adults + 1 child**, so the returned CZK price is treated as the total party price, not a per-person price.

The current alert_price_czk is an absolute emergency threshold. The normal alert mechanism is the price drop compared with the previous best.

## Outputs

- data/prices.csv - raw historical flight observations
- data/flight_prices.xlsx - dashboard, history and complete 6–8 night combinations
- data/alert_state.json - last alert baseline

## GitHub Actions

The workflow runs three times per week and can also be started manually from GitHub Actions.

Required repository secret:

SERPAPI_KEY

The workflow needs repository permission to create Issues because price alerts are delivered as GitHub Issues.

## Local test

~~~powershell
python src/test_google_flights.py
~~~

Data source: Google Flights through SerpApi.
