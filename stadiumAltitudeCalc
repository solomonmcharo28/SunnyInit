import pandas as pd
import requests
from datetime import date
from statistics import mean
import time

INPUT_XLSX = "StadiumAltitude.xlsx"
OUTPUT_XLSX = "StadiumAltitude_with_weather.xlsx"

# --- CONFIG ---
# "Last one year" as [today - 365 days, today] (inclusive end_date)
END_DATE = date.today()
START_DATE = date(END_DATE.year - 1, END_DATE.month, END_DATE.day)

DAILY_VARS = [
    "temperature_2m_mean",
    "relative_humidity_2m_mean",
    "pressure_msl_mean",
    "cloud_cover_mean",
    "precipitation_sum",
]

API_URL = "https://archive-api.open-meteo.com/v1/archive"

def fetch_daily(lat: float, lon: float, start_date: str, end_date: str) -> dict:
    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": start_date,
        "end_date": end_date,
        "daily": ",".join(DAILY_VARS),
        "timezone": "auto",
    }
    r = requests.get(API_URL, params=params, timeout=30)
    r.raise_for_status()
    j = r.json()
    if "daily" not in j:
        raise ValueError(f"No 'daily' field in response for lat={lat}, lon={lon}: {j}")
    return j["daily"]

def annual_facts_from_daily(daily: dict) -> dict:
    # Defensive: make sure lists exist
    t = daily.get("temperature_2m_mean", [])
    rh = daily.get("relative_humidity_2m_mean", [])
    p = daily.get("pressure_msl_mean", [])
    cc = daily.get("cloud_cover_mean", [])
    pr = daily.get("precipitation_sum", [])

    # Some locations/dates may have missing values (None) – filter them out
    def clean(xs):
        return [x for x in xs if x is not None]

    t, rh, p, cc, pr = map(clean, (t, rh, p, cc, pr))

    return {
        "annual_avg_temp_c": mean(t) if t else None,
        "annual_avg_humidity_pct": mean(rh) if rh else None,
        "annual_avg_pressure_hpa": mean(p) if p else None,
        "annual_avg_cloudcover_pct": mean(cc) if cc else None,
        "annual_total_precip_mm": sum(pr) if pr else None,
    }

def main():
    df = pd.read_excel(INPUT_XLSX)

    # Normalize column names (handles your "Latitude " trailing space)
    rename_map = {c: c.strip() for c in df.columns}
    df.rename(columns=rename_map, inplace=True)

    if "Latitude" not in df.columns or "Longitude" not in df.columns:
        raise KeyError("Could not find 'Latitude' and 'Longitude' columns after stripping whitespace.")

    start_s = START_DATE.isoformat()
    end_s = END_DATE.isoformat()

    # Add columns if missing
    new_cols = [
        "annual_avg_temp_c",
        "annual_avg_humidity_pct",
        "annual_avg_pressure_hpa",
        "annual_avg_cloudcover_pct",
        "annual_total_precip_mm",
        "weather_start_date",
        "weather_end_date",
    ]
    for c in new_cols:
        if c not in df.columns:
            df[c] = None

    # Simple cache to avoid duplicate calls if coords repeat
    cache = {}

    for i, row in df.iterrows():
        lat = row["Latitude"]
        lon = row["Longitude"]

        if pd.isna(lat) or pd.isna(lon):
            continue

        key = (round(float(lat), 6), round(float(lon), 6), start_s, end_s)
        if key in cache:
            facts = cache[key]
        else:
            daily = fetch_daily(float(lat), float(lon), start_s, end_s)
            facts = annual_facts_from_daily(daily)
            cache[key] = facts
            time.sleep(0.2)  # polite pacing

        for k, v in facts.items():
            df.at[i, k] = v

        df.at[i, "weather_start_date"] = start_s
        df.at[i, "weather_end_date"] = end_s

    df.to_excel(OUTPUT_XLSX, index=False)
    print(f"Saved: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
