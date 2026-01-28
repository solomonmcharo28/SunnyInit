#!/usr/bin/env python3
"""
Enrich an Excel sheet of players with:
- Country of league_name (via Wikidata SPARQL)
- League-country capital elevation, mean humidity, and mean temperature (Open-Meteo)
- FUTBIN price on New Year's Day for EA FC 25 and EA FC 26 (via playerGraph JSON)
Hint - Querying the futbin playergraph is hard due to a mismatch between the EA Sports FC Player ID and the FUTBin ID - Solomon Mcharo, Constance Develle, Britanny Quan, Jiaqi Paige, Claudia Sinclair and Madeline Young

Requires: pandas, requests, openpyxl
pip install pandas requests openpyxl
"""

from __future__ import annotations

import argparse
import json
import math
import time
from dataclasses import dataclass
from datetime import datetime, date
from typing import Any, Dict, Optional, Tuple, List

import pandas as pd
import requests


# -----------------------------
# Config
# -----------------------------
WIKIDATA_SPARQL_URL = "https://query.wikidata.org/sparql"

OPEN_METEO_ELEVATION_URL = "https://api.open-meteo.com/v1/elevation"
OPEN_METEO_CLIMATE_URL = "https://climate-api.open-meteo.com/v1/climate"

# FUTBIN graph endpoint pattern (commonly used by scrapers)
# Example pattern documented in scraping Q/A: https://www.futbin.com/19/playerGraph?type=daily_graph&year=19&player=...
# We'll use /{year}/playerGraph for modern years.
FUTBIN_PLAYER_GRAPH_URL = "https://www.futbin.com/{year}/playerGraph"

DEFAULT_PLATFORM_KEY = "ps"  # keys in FUTBIN graph JSON often: ps, xbox, pc

# Choose what "New Year's Day" means for each game
FC25_NYD = date(2025, 1, 1)
FC26_NYD = date(2026, 1, 1)

# Open-Meteo climate averaging window: last full calendar year before "now"
# You can change this to match your research needs.
CLIMATE_START = "2024-01-01"
CLIMATE_END = "2024-12-31"

REQUEST_TIMEOUT = 30
SLEEP_BETWEEN_CALLS_SEC = 0.2  # basic politeness to reduce rate-limit risk


# -----------------------------
# Helpers
# -----------------------------
def _sleep():
    time.sleep(SLEEP_BETWEEN_CALLS_SEC)


def http_get_json(url: str, params: Dict[str, Any] | None = None, headers: Dict[str, str] | None = None) -> Any:
    r = requests.get(url, params=params, headers=headers, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    # Some endpoints may return text/html unexpectedly if blocked; handle gracefully
    ctype = r.headers.get("Content-Type", "")
    if "application/json" in ctype or r.text.strip().startswith("{") or r.text.strip().startswith("["):
        return r.json()
    raise ValueError(f"Expected JSON but got Content-Type={ctype} for {url}")


def http_post_sparql(query: str) -> Any:
    headers = {
        "Accept": "application/sparql-results+json",
        "User-Agent": "league-country-enricher/1.0 (research script; contact: none)",
    }
    data = {"query": query}
    r = requests.post(WIKIDATA_SPARQL_URL, data=data, headers=headers, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    return r.json()


def safe_float(x: Any) -> Optional[float]:
    try:
        if x is None:
            return None
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).strip()
        if s == "" or s.lower() in {"nan", "none", "null"}:
            return None
        return float(s)
    except Exception:
        return None


def closest_date_price(series: List[List[Any]], target: date) -> Optional[int]:
    """
    series is typically [[unix_ms, price], ...]
    Return the price on the closest date to target (by absolute day difference).
    """
    if not series:
        return None

    best = None
    best_delta = None

    for item in series:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue
        ts_ms, price = item[0], item[1]
        ts = datetime.utcfromtimestamp(ts_ms / 1000.0).date()
        delta = abs((ts - target).days)
        if best is None or delta < best_delta:
            best = price
            best_delta = delta

    if best is None:
        return None

    try:
        return int(best)
    except Exception:
        return None


# -----------------------------
# Wikidata league -> country (+ capital coords)
# -----------------------------
@dataclass(frozen=True)
class LeagueGeo:
    league_country: Optional[str]
    capital_name: Optional[str]
    capital_lat: Optional[float]
    capital_lon: Optional[float]


def wikidata_league_to_country_and_capital(league_name: str) -> LeagueGeo:
    """
    Best-effort: find an association football league with this English label,
    then return country and the country's capital coordinates.

    Wikidata properties used:
    - P17: country
    - P36: capital
    - P625: coordinate location
    """
    # Try to match: (league) instance of association football league (Q15991303) OR sports league (Q623109)
    query = f"""
    SELECT ?countryLabel ?capitalLabel ?lat ?lon WHERE {{
      ?league rdfs:label "{league_name}"@en .
      ?league wdt:P31/wdt:P279* wd:Q15991303 .
      OPTIONAL {{ ?league wdt:P17 ?country . }}
      OPTIONAL {{ ?league wdt:P495 ?country . }}  # sometimes "country of origin"
      OPTIONAL {{ ?country wdt:P36 ?capital . }}
      OPTIONAL {{
        ?capital wdt:P625 ?coord .
        BIND(geof:latitude(?coord) as ?lat)
        BIND(geof:longitude(?coord) as ?lon)
      }}
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
    }}
    LIMIT 5
    """
    js = http_post_sparql(query)
    _sleep()

    bindings = js.get("results", {}).get("bindings", [])
    if not bindings:
        return LeagueGeo(None, None, None, None)

    b = bindings[0]
    country = b.get("countryLabel", {}).get("value")
    capital = b.get("capitalLabel", {}).get("value")
    lat = safe_float(b.get("lat", {}).get("value"))
    lon = safe_float(b.get("lon", {}).get("value"))
    return LeagueGeo(country, capital, lat, lon)


# -----------------------------
# Open-Meteo elevation + climate
# -----------------------------
@dataclass(frozen=True)
class ClimateSummary:
    elevation_m: Optional[float]
    mean_humidity_2m: Optional[float]
    mean_temp_2m_c: Optional[float]


def open_meteo_elevation(lat: float, lon: float) -> Optional[float]:
    js = http_get_json(OPEN_METEO_ELEVATION_URL, params={"latitude": lat, "longitude": lon})
    _sleep()
    elev = js.get("elevation")
    if isinstance(elev, list) and elev:
        return safe_float(elev[0])
    return None


def open_meteo_climate_means(lat: float, lon: float, start: str, end: str) -> Tuple[Optional[float], Optional[float]]:
    """
    Use Open-Meteo Climate API daily time series, then average:
    - relative_humidity_2m_mean
    - temperature_2m_mean (used here as a simple climate proxy)
    """
    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": start,
        "end_date": end,
        "models": "ERA5",
        "daily": "relative_humidity_2m_mean,temperature_2m_mean",
        "timezone": "UTC",
    }
    js = http_get_json(OPEN_METEO_CLIMATE_URL, params=params)
    _sleep()

    daily = js.get("daily", {})
    hum = daily.get("relative_humidity_2m_mean", [])
    tmp = daily.get("temperature_2m_mean", [])

    def mean(xs):
        vals = [safe_float(x) for x in xs]
        vals = [v for v in vals if v is not None and not math.isnan(v)]
        return float(sum(vals) / len(vals)) if vals else None

    return mean(hum), mean(tmp)


def climate_summary_for_capital(lat: Optional[float], lon: Optional[float]) -> ClimateSummary:
    if lat is None or lon is None:
        return ClimateSummary(None, None, None)

    elev = open_meteo_elevation(lat, lon)
    hum_mean, temp_mean = open_meteo_climate_means(lat, lon, CLIMATE_START, CLIMATE_END)
    return ClimateSummary(elev, hum_mean, temp_mean)


# -----------------------------
# FUTBIN pricing
# -----------------------------
def futbin_player_graph(year: int, player_id: int) -> Dict[str, Any]:
    """
    Fetch FUTBIN player graph JSON for the given year + player id.
    """
    url = FUTBIN_PLAYER_GRAPH_URL.format(year=year)
    params = {
        "type": "daily_graph",
        "year": year,
        "player": player_id,
        "set_id": "",  # usually present, often empty
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; research script)",
        "Accept": "application/json,text/plain,*/*",
        "Referer": f"https://www.futbin.com/{year}/player/{player_id}",
    }
    js = http_get_json(url, params=params, headers=headers)
    _sleep()
    return js


def futbin_price_on_date(year: int, player_id: int, target_day: date, platform_key: str = DEFAULT_PLATFORM_KEY) -> Optional[int]:
    """
    Returns price on target_day (or nearest day in series).
    """
    try:
        js = futbin_player_graph(year, player_id)
    except Exception:
        return None

    series = js.get(platform_key)
    if not isinstance(series, list):
        return None

    return closest_date_price(series, target_day)


def futbin_search_best_effort(year: int, term: str) -> Optional[int]:
    """
    Best-effort attempt to resolve a FUTBIN player id via FUTBIN search.
    This is NOT guaranteed stable. If it fails, provide futbin_player_id in your sheet.

    Many community scripts use /search?year=YY&term=...
    """
    url = "https://www.futbin.com/search"
    params = {"year": year, "term": term}
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; research script)",
        "Accept": "application/json,text/plain,*/*",
        "Referer": f"https://www.futbin.com/{year}/players",
    }
    try:
        js = http_get_json(url, params=params, headers=headers)
        _sleep()
    except Exception:
        return None

    # Try common shapes:
    # - list of dicts
    # - {"players":[...]}
    items = None
    if isinstance(js, list):
        items = js
    elif isinstance(js, dict):
        for k in ("players", "items", "results"):
            if isinstance(js.get(k), list):
                items = js[k]
                break

    if not items:
        return None

    # Heuristic: pick first item with an id-like field
    for it in items[:10]:
        if not isinstance(it, dict):
            continue
        for key in ("player_id", "id", "resource_id"):
            v = it.get(key)
            if isinstance(v, int):
                return v
            if isinstance(v, str) and v.isdigit():
                return int(v)

    return None


# -----------------------------
# Main pipeline
# -----------------------------
def enrich(input_xlsx: str, output_xlsx: str, platform_key: str = DEFAULT_PLATFORM_KEY) -> None:
    df = pd.read_excel(input_xlsx)

    # Caches to avoid repeated API calls
    league_geo_cache: Dict[str, LeagueGeo] = {}
    climate_cache: Dict[Tuple[Optional[float], Optional[float]], ClimateSummary] = {}

    # Ensure output columns exist
    for col in [
        "league_country",
        "country_capital",
        "capital_lat",
        "capital_lon",
        "avg_altitude_m",
        "avg_humidity_2m",
        "avg_temp_2m_c",
        "futbin_player_id_25",
        "futbin_player_id_26",
        "price_fc25_nyd",
        "price_fc26_nyd",
    ]:
        if col not in df.columns:
            df[col] = None

    # Prefer user-provided FUTBIN ids if present
    has_futbin_player_id = "futbin_player_id" in df.columns

    for i, row in df.iterrows():
        league_name = str(row.get("league_name", "")).strip()

        # 1) league -> country (+ capital coords)
        if league_name:
            if league_name not in league_geo_cache:
                try:
                    league_geo_cache[league_name] = wikidata_league_to_country_and_capital(league_name)
                except Exception:
                    league_geo_cache[league_name] = LeagueGeo(None, None, None, None)

            geo = league_geo_cache[league_name]
            df.at[i, "league_country"] = geo.league_country
            df.at[i, "country_capital"] = geo.capital_name
            df.at[i, "capital_lat"] = geo.capital_lat
            df.at[i, "capital_lon"] = geo.capital_lon

            # 2) climate summary for capital (proxy for the league country)
            key = (geo.capital_lat, geo.capital_lon)
            if key not in climate_cache:
                try:
                    climate_cache[key] = climate_summary_for_capital(geo.capital_lat, geo.capital_lon)
                except Exception:
                    climate_cache[key] = ClimateSummary(None, None, None)

            cs = climate_cache[key]
            df.at[i, "avg_altitude_m"] = cs.elevation_m
            df.at[i, "avg_humidity_2m"] = cs.mean_humidity_2m
            df.at[i, "avg_temp_2m_c"] = cs.mean_temp_2m_c

        # 3) FUTBIN player id resolution
        name = str(row.get("full_name", "")).strip()
        futbin_id = None
        if has_futbin_player_id:
            v = row.get("futbin_player_id")
            if isinstance(v, (int, float)) and not pd.isna(v):
                futbin_id = int(v)
            elif isinstance(v, str) and v.strip().isdigit():
                futbin_id = int(v.strip())

        # If no id provided, try best-effort search for each year
        # (ids may differ across years; we store per-year ids)
        futbin_id_25 = None
        futbin_id_26 = None

        if futbin_id is not None:
            # assume same id works for both years (not always true)
            futbin_id_25 = futbin_id
            futbin_id_26 = futbin_id
        else:
            if name:
                futbin_id_25 = futbin_search_best_effort(25, name)
                futbin_id_26 = futbin_search_best_effort(26, name)

        df.at[i, "futbin_player_id_25"] = futbin_id_25
        df.at[i, "futbin_player_id_26"] = futbin_id_26

        # 4) prices on NYD (or nearest available day)
        if futbin_id_25 is not None:
            df.at[i, "price_fc25_nyd"] = futbin_price_on_date(25, int(futbin_id_25), FC25_NYD, platform_key=platform_key)

        if futbin_id_26 is not None:
            df.at[i, "price_fc26_nyd"] = futbin_price_on_date(26, int(futbin_id_26), FC26_NYD, platform_key=platform_key)

    # Write to new workbook
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="enriched")

    print(f"Saved enriched workbook -> {output_xlsx}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Input .xlsx path")
    ap.add_argument("--output", required=True, help="Output .xlsx path")
    ap.add_argument("--platform", default=DEFAULT_PLATFORM_KEY, choices=["ps", "xbox", "pc"],
                    help="Which FUTBIN platform price series to use")
    args = ap.parse_args()

    enrich(args.input, args.output, platform_key=args.platform)


if __name__ == "__main__":
    main()
