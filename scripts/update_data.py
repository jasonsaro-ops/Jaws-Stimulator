#!/usr/bin/env python3
"""
JAWS STIMULATOR — Daily GSAF data updater
Downloads the latest Global Shark Attack File (Excel), cleans it,
adds improved coastal / state / location coordinates, and writes data/shark_data.json.
"""

from __future__ import annotations

import json
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests

GSAF_URL = "https://www.sharkattackfile.net/spreadsheets/GSAF5.xls"
ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "data" / "shark_data.json"
META_JSON = ROOT / "data" / "meta.json"
TMP_XLS = ROOT / "data" / "_gsaf_tmp.xls"
MAX_RECORDS = 8000
MIN_YEAR = 1800
RANDOM_SEED = 42

COUNTRY_CENTROIDS = {
    "USA": (27.8, -81.5),
    "UNITED STATES": (27.8, -81.5),
    "AUSTRALIA": (-25.0, 135.0),
    "SOUTH AFRICA": (-30.5, 25.0),
    "PAPUA NEW GUINEA": (-6.0, 147.0),
    "BAHAMAS": (24.2, -76.0),
    "BRAZIL": (-15.0, -45.0),
    "NEW ZEALAND": (-40.0, 175.0),
    "MEXICO": (20.5, -100.0),
    "NEW CALEDONIA": (-21.3, 165.5),
    "ITALY": (41.0, 12.5),
    "FIJI": (-17.8, 178.0),
    "REUNION": (-21.1, 55.5),
    "EGYPT": (27.0, 33.5),
    "MOZAMBIQUE": (-18.5, 35.5),
    "PHILIPPINES": (12.0, 122.0),
    "JAPAN": (34.5, 135.0),
    "INDIA": (15.0, 75.0),
    "FRANCE": (46.0, 2.0),
    "SPAIN": (40.0, -3.5),
    "UNITED KINGDOM": (54.0, -2.0),
    "CANADA": (45.0, -75.0),
    "CUBA": (22.0, -80.0),
    "COSTA RICA": (9.8, -84.0),
    "FRENCH POLYNESIA": (-17.5, -149.5),
    "SOLOMON ISLANDS": (-9.0, 160.0),
    "TONGA": (-21.0, -175.0),
    "VANUATU": (-16.0, 167.0),
    "INDONESIA": (-2.0, 118.0),
    "THAILAND": (13.5, 100.5),
    "VIETNAM": (16.0, 108.0),
    "SRI LANKA": (7.0, 80.0),
    "IRAN": (27.0, 53.0),
    "TURKEY": (38.5, 35.0),
    "HONG KONG": (22.3, 114.2),
    "SENEGAL": (14.5, -17.0),
    "KENYA": (-3.0, 40.0),
    "MAURITIUS": (-20.3, 57.5),
    "SEYCHELLES": (-4.6, 55.5),
    "JAMAICA": (18.1, -77.3),
    "PANAMA": (8.5, -80.0),
    "BERMUDA": (32.3, -64.8),
    "ECUADOR": (-1.0, -80.0),
    "COLOMBIA": (10.0, -75.0),
    "VENEZUELA": (10.5, -66.5),
    "CHILE": (-33.0, -71.5),
    "ARGENTINA": (-38.0, -57.5),
    "PERU": (-12.0, -77.0),
    "MALDIVES": (3.2, 73.2),
    "SAMOA": (-13.8, -171.8),
    "AMERICAN SAMOA": (-14.3, -170.7),
    "GUAM": (13.4, 144.8),
}

US_STATE_CENTROIDS = {
    "FLORIDA": (27.5, -81.5),
    "HAWAII": (20.8, -156.5),
    "CALIFORNIA": (35.5, -121.0),
    "SOUTH CAROLINA": (32.8, -79.8),
    "NORTH CAROLINA": (34.5, -76.5),
    "TEXAS": (28.0, -96.0),
    "NEW JERSEY": (39.5, -74.2),
    "NEW YORK": (40.6, -73.5),
    "OREGON": (44.5, -124.0),
    "MASSACHUSETTS": (41.8, -70.5),
    "VIRGINIA": (37.0, -76.0),
    "GEORGIA": (31.5, -81.0),
    "LOUISIANA": (29.5, -90.0),
    "ALABAMA": (30.3, -87.8),
    "DELAWARE": (38.78, -75.10),
    "MARYLAND": (38.3, -75.1),
    "CONNECTICUT": (41.2, -72.8),
    "RHODE ISLAND": (41.4, -71.5),
    "MAINE": (43.7, -69.8),
    "WASHINGTON": (47.5, -124.0),
    "ALASKA": (58.0, -135.0),
    "PUERTO RICO": (18.2, -66.5),
    "US VIRGIN ISLANDS": (18.3, -64.8),
    "VIRGIN ISLANDS": (18.3, -64.8),
}

LOCATION_OVERRIDES = [
    ("HERRING POINT", 38.765, -75.082),
    ("CAPE HENLOPEN", 38.787, -75.094),
    ("DEWEY BEACH", 38.692, -75.075),
    ("INDIAN RIVER INLET", 38.610, -75.068),
    ("FENWICK ISLAND", 38.460, -75.050),
    ("MISPILLION", 38.947, -75.315),
    ("LEWES", 38.775, -75.140),
    ("REHOBOTH", 38.720, -75.080),
    ("DAYTONA", 29.210, -81.020),
    ("NEW SMYRNA", 29.025, -80.925),
    ("VOLUSIA", 29.15, -81.00),
    ("MURRELLS INLET", 33.55, -79.03),
    ("MYRTLE BEACH", 33.69, -78.89),
    ("HUNTINGTON BEACH", 33.66, -118.00),
    ("BONDI", -33.89, 151.27),
    ("SYDNEY", -33.87, 151.21),
    ("PERTH", -31.95, 115.86),
    ("DURBAN", -29.86, 31.03),
    ("FALSE BAY", -34.20, 18.65),
    ("REUNION", -21.12, 55.53),
    ("RECIFE", -8.05, -34.88),
]


def download_gsaf(url: str, dest: Path) -> None:
    print(f"Downloading GSAF from {url} …")
    r = requests.get(
        url,
        timeout=120,
        headers={"User-Agent": "JawsStimulator/1.1 (GitHub Action; educational dashboard)"},
    )
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)
    print(f"  Saved {len(r.content):,} bytes → {dest}")


def normalize_country(c: str) -> str:
    if not c or not isinstance(c, str):
        return "UNKNOWN"
    c = c.upper().strip()
    aliases = {
        "UNITED STATES": "USA",
        "U.S.A.": "USA",
        "U.S.": "USA",
        "USA": "USA",
        "ENGLAND": "UNITED KINGDOM",
        "SCOTLAND": "UNITED KINGDOM",
        "WALES": "UNITED KINGDOM",
        "REUNION ISLAND": "REUNION",
        "RÉUNION": "REUNION",
    }
    return aliases.get(c, c)


def get_coords(country: str, state: str, location: str) -> tuple[float, float, str]:
    loc_u = (location or "").upper()
    state_u = (state or "").upper()
    country_n = normalize_country(country)

    for key, lat, lon in LOCATION_OVERRIDES:
        if key in loc_u or key in state_u:
            return lat, lon, f"location:{key}"

    if country_n == "USA":
        for st, (lat, lon) in US_STATE_CENTROIDS.items():
            if st in state_u or state_u == st:
                return lat, lon, f"state:{st}"
        if any(x in loc_u for x in ("HAWAII", "OAHU", "MAUI", "KAUAI")):
            return (*US_STATE_CENTROIDS["HAWAII"], "state:HAWAII")

    if country_n in COUNTRY_CENTROIDS:
        lat, lon = COUNTRY_CENTROIDS[country_n]
        return lat, lon, f"country:{country_n}"

    for k, v in COUNTRY_CENTROIDS.items():
        if k in country_n or country_n in k:
            return v[0], v[1], f"country:{k}"

    return 0.0, 0.0, "unknown"


def process(xls_path: Path) -> list[dict]:
    print("Reading Excel …")
    try:
        df = pd.read_excel(xls_path, engine="xlrd")
    except Exception:
        df = pd.read_excel(xls_path, engine="openpyxl")

    df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(df.columns)]
    print(f"  Columns: {list(df.columns)[:14]}…")
    print(f"  Raw rows: {len(df)}")

    if "Year" not in df.columns:
        raise RuntimeError("No 'Year' column found in GSAF file")

    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df = df.dropna(subset=["Year"])
    df = df[(df["Year"] >= MIN_YEAR) & (df["Year"] <= datetime.now().year + 1)]
    df = df.sort_values("Year", ascending=False)

    if len(df) > MAX_RECORDS:
        df = df.head(MAX_RECORDS)
    print(f"  After year filter ({MIN_YEAR}+) & limit: {len(df)} rows")

    random.seed(RANDOM_SEED)
    records = []
    for _, row in df.iterrows():
        country_raw = str(row.get("Country", "") or "").strip()
        state = str(row.get("State", "") or "").strip() if pd.notna(row.get("State")) else ""
        location = str(row.get("Location", "") or "").strip() if pd.notna(row.get("Location")) else ""

        base_lat, base_lon, coord_src = get_coords(country_raw, state, location)

        if coord_src.startswith("location:"):
            lat, lon = base_lat, base_lon
        elif coord_src.startswith("state:"):
            lat = base_lat + random.uniform(-0.35, 0.35)
            lon = base_lon + random.uniform(-0.45, 0.45)
        else:
            lat = base_lat + random.uniform(-2.5, 2.5)
            lon = base_lon + random.uniform(-3.5, 3.5)

        lat = max(-85.0, min(85.0, lat))
        lon = max(-180.0, min(180.0, lon))

        fatal_raw = str(row.get("Fatal Y/N", "") or "").upper().strip()
        fatal = fatal_raw.startswith("Y")

        species = str(row.get("Species", "") or row.get("Species ", "") or "").strip()
        if species.lower() in ("nan", "none", ""):
            species = ""

        attack_type = str(row.get("Type", "") or "").strip()

        rec = {
            "id": len(records) + 1,
            "date": str(row.get("Date", "") or "").strip(),
            "year": int(row["Year"]),
            "type": attack_type,
            "country": country_raw or normalize_country(country_raw),
            "state": state,
            "location": location,
            "activity": str(row.get("Activity", "") or "").strip() if pd.notna(row.get("Activity")) else "",
            "name": str(row.get("Name", "") or "").strip() if pd.notna(row.get("Name")) else "",
            "sex": str(row.get("Sex", "") or "").strip() if pd.notna(row.get("Sex")) else "",
            "age": str(row.get("Age", "") or "").strip() if pd.notna(row.get("Age")) else "",
            "injury": str(row.get("Injury", "") or "").strip() if pd.notna(row.get("Injury")) else "",
            "fatal": fatal,
            "time": str(row.get("Time", "") or "").strip() if pd.notna(row.get("Time")) else "",
            "species": species[:90],
            "source": str(row.get("Source", "") or "").strip()[:150] if pd.notna(row.get("Source")) else "",
            "lat": round(lat, 4),
            "lon": round(lon, 4),
        }
        records.append(rec)

    hp = [r for r in records if "HERRING" in (r["location"] or "").upper()]
    if hp:
        print(f"  Herring Point sample: lat={hp[0]['lat']}, lon={hp[0]['lon']} (expect ~38.76, -75.08)")

    return records


def main() -> int:
    print("=" * 60)
    print("JAWS STIMULATOR — GSAF data update")
    print(f"UTC time : {datetime.now(timezone.utc).isoformat()}")
    print("=" * 60)

    try:
        download_gsaf(GSAF_URL, TMP_XLS)
        records = process(TMP_XLS)

        OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False)

        years = [r["year"] for r in records if r.get("year")]
        meta = {
            "updated": datetime.now(timezone.utc).isoformat(),
            "source": "Global Shark Attack File (GSAF) — https://www.sharkattackfile.net",
            "note": "Coordinates use location overrides → US state centroids → country coastal points.",
            "count": len(records),
            "min_year": min(years) if years else MIN_YEAR,
            "max_year": max(years) if years else None,
            "max_records": MAX_RECORDS,
        }
        with open(META_JSON, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)

        print(f"Wrote {len(records)} records → {OUT_JSON}")
        print(f"Meta → {META_JSON}")
        print(f"Year range: {meta['min_year']} – {meta['max_year']}")

        if TMP_XLS.exists():
            TMP_XLS.unlink()

        print("Done.")
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
