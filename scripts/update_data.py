#!/usr/bin/env python3
"""
JAWS STIMULATOR — Multi-source shark incident data updater
Sources:
  - Global Shark Attack File (GSAF) — primary global registry
  - California CDFW Shark Incident Database (with real lat/lon)
Records are tagged with a `source` field for filtering.
"""

from __future__ import annotations

import json
import random
import sys
import zipfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests

GSAF_URL = "https://www.sharkattackfile.net/spreadsheets/GSAF5.xls"
CA_ZIP_URL = (
    "https://data.cnra.ca.gov/dataset/022f670f-eb4e-4314-9fad-74f7f9560400/"
    "resource/fa77215c-a06c-4fa2-b8bc-ffbcaaa65286/download/sharkincidents_1950_2022_220302.zip"
)

ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "data" / "shark_data.json"
META_JSON = ROOT / "data" / "meta.json"
TMP_XLS = ROOT / "data" / "_gsaf_tmp.xls"
MAX_GSAF = 8000
MIN_YEAR = 1800
RANDOM_SEED = 42

COUNTRY_CENTROIDS = {
    "USA": (27.8, -81.5), "UNITED STATES": (27.8, -81.5),
    "AUSTRALIA": (-25.0, 135.0), "SOUTH AFRICA": (-30.5, 25.0),
    "PAPUA NEW GUINEA": (-6.0, 147.0), "BAHAMAS": (24.2, -76.0),
    "BRAZIL": (-15.0, -45.0), "NEW ZEALAND": (-40.0, 175.0),
    "MEXICO": (20.5, -100.0), "NEW CALEDONIA": (-21.3, 165.5),
    "ITALY": (41.0, 12.5), "FIJI": (-17.8, 178.0), "REUNION": (-21.1, 55.5),
    "EGYPT": (27.0, 33.5), "MOZAMBIQUE": (-18.5, 35.5), "PHILIPPINES": (12.0, 122.0),
    "JAPAN": (34.5, 135.0), "INDIA": (15.0, 75.0), "FRANCE": (46.0, 2.0),
    "SPAIN": (40.0, -3.5), "UNITED KINGDOM": (54.0, -2.0), "CANADA": (45.0, -75.0),
    "CUBA": (22.0, -80.0), "COSTA RICA": (9.8, -84.0),
    "FRENCH POLYNESIA": (-17.5, -149.5), "SOLOMON ISLANDS": (-9.0, 160.0),
    "TONGA": (-21.0, -175.0), "VANUATU": (-16.0, 167.0), "INDONESIA": (-2.0, 118.0),
    "THAILAND": (13.5, 100.5), "VIETNAM": (16.0, 108.0), "SRI LANKA": (7.0, 80.0),
    "IRAN": (27.0, 53.0), "TURKEY": (38.5, 35.0), "HONG KONG": (22.3, 114.2),
    "SENEGAL": (14.5, -17.0), "KENYA": (-3.0, 40.0), "MAURITIUS": (-20.3, 57.5),
    "SEYCHELLES": (-4.6, 55.5), "JAMAICA": (18.1, -77.3), "PANAMA": (8.5, -80.0),
    "BERMUDA": (32.3, -64.8), "ECUADOR": (-1.0, -80.0), "COLOMBIA": (10.0, -75.0),
    "VENEZUELA": (10.5, -66.5), "CHILE": (-33.0, -71.5), "ARGENTINA": (-38.0, -57.5),
    "PERU": (-12.0, -77.0), "MALDIVES": (3.2, 73.2), "SAMOA": (-13.8, -171.8),
    "GUAM": (13.4, 144.8),
}

US_STATE_CENTROIDS = {
    "FLORIDA": (27.5, -81.5), "HAWAII": (20.8, -156.5), "CALIFORNIA": (36.5, -121.5),
    "SOUTH CAROLINA": (32.8, -79.8), "NORTH CAROLINA": (34.5, -76.5),
    "TEXAS": (28.0, -96.0), "NEW JERSEY": (39.5, -74.2), "NEW YORK": (40.6, -73.5),
    "OREGON": (44.5, -124.0), "MASSACHUSETTS": (41.8, -70.5), "VIRGINIA": (37.0, -76.0),
    "GEORGIA": (31.5, -81.0), "LOUISIANA": (29.5, -90.0), "ALABAMA": (30.3, -87.8),
    "DELAWARE": (38.78, -75.10), "MARYLAND": (38.3, -75.1), "CONNECTICUT": (41.2, -72.8),
    "RHODE ISLAND": (41.4, -71.5), "MAINE": (43.7, -69.8), "WASHINGTON": (47.5, -124.0),
    "ALASKA": (58.0, -135.0), "PUERTO RICO": (18.2, -66.5),
}

LOCATION_OVERRIDES = [
    ("HERRING POINT", 38.765, -75.082), ("CAPE HENLOPEN", 38.787, -75.094),
    ("DEWEY BEACH", 38.692, -75.075), ("INDIAN RIVER INLET", 38.610, -75.068),
    ("FENWICK ISLAND", 38.460, -75.050), ("MISPILLION", 38.947, -75.315),
    ("LEWES", 38.775, -75.140), ("REHOBOTH", 38.720, -75.080),
    ("DAYTONA", 29.210, -81.020), ("NEW SMYRNA", 29.025, -80.925),
    ("VOLUSIA", 29.15, -81.00), ("MURRELLS INLET", 33.55, -79.03),
    ("MYRTLE BEACH", 33.69, -78.89), ("HUNTINGTON BEACH", 33.66, -118.00),
    ("BONDI", -33.89, 151.27), ("SYDNEY", -33.87, 151.21), ("PERTH", -31.95, 115.86),
    ("DURBAN", -29.86, 31.03), ("FALSE BAY", -34.20, 18.65),
    ("REUNION", -21.12, 55.53), ("RECIFE", -8.05, -34.88),
    ("MAKAHA", 21.469, -158.220), ("PIPELINE", 21.665, -158.053),
    ("NORTH SHORE", 21.68, -158.05), ("WAIKIKI", 21.276, -157.827),
]


def normalize_country(c: str) -> str:
    if not c or not isinstance(c, str):
        return "UNKNOWN"
    c = c.upper().strip()
    aliases = {
        "UNITED STATES": "USA", "U.S.A.": "USA", "U.S.": "USA", "USA": "USA",
        "ENGLAND": "UNITED KINGDOM", "SCOTLAND": "UNITED KINGDOM", "WALES": "UNITED KINGDOM",
        "REUNION ISLAND": "REUNION", "RÉUNION": "REUNION",
    }
    return aliases.get(c, c)


def get_coords(country: str, state: str, location: str) -> tuple[float, float, str]:
    loc_u = (location or "").upper()
    state_u = (state or "").upper()
    country_n = normalize_country(country)

    for key, lat, lon in LOCATION_OVERRIDES:
        if key in loc_u:
            return lat, lon, f"location:{key}"

    if country_n == "USA":
        for st, (lat, lon) in US_STATE_CENTROIDS.items():
            if st in state_u or state_u == st:
                return lat, lon, f"state:{st}"
        if any(x in loc_u for x in ("HAWAII", "OAHU", "MAUI", "KAUAI", "BIG ISLAND")):
            return (*US_STATE_CENTROIDS["HAWAII"], "state:HAWAII")

    if country_n in COUNTRY_CENTROIDS:
        lat, lon = COUNTRY_CENTROIDS[country_n]
        return lat, lon, f"country:{country_n}"
    for k, v in COUNTRY_CENTROIDS.items():
        if k in country_n or country_n in k:
            return v[0], v[1], f"country:{k}"
    return 0.0, 0.0, "unknown"


def download(url: str, timeout: int = 120) -> bytes:
    r = requests.get(url, timeout=timeout, headers={
        "User-Agent": "JawsStimulator/2.0 (GitHub Action; educational multi-source dashboard)"
    })
    r.raise_for_status()
    return r.content


def process_gsaf(xls_bytes: bytes) -> list[dict]:
    print("Processing GSAF …")
    df = pd.read_excel(BytesIO(xls_bytes), engine="xlrd")
    df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(df.columns)]
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df = df.dropna(subset=["Year"])
    df = df[(df["Year"] >= MIN_YEAR) & (df["Year"] <= datetime.now().year + 1)]
    df = df.sort_values("Year", ascending=False)
    if len(df) > MAX_GSAF:
        df = df.head(MAX_GSAF)
    print(f"  GSAF rows kept: {len(df)}")

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
            lat = base_lat + random.uniform(-0.25, 0.25)
            lon = base_lon + random.uniform(-0.35, 0.35)
        else:
            lat = base_lat + random.uniform(-2.0, 2.0)
            lon = base_lon + random.uniform(-3.0, 3.0)
        lat = max(-85.0, min(85.0, lat))
        lon = max(-180.0, min(180.0, lon))

        fatal_raw = str(row.get("Fatal Y/N", "") or "").upper().strip()
        species = str(row.get("Species", "") or "").strip()
        if species.lower() in ("nan", "none", ""):
            species = ""

        records.append({
            "id": f"gsaf-{len(records)+1}",
            "date": str(row.get("Date", "") or "").strip(),
            "year": int(row["Year"]),
            "type": str(row.get("Type", "") or "").strip(),
            "country": country_raw or normalize_country(country_raw),
            "state": state,
            "location": location,
            "activity": str(row.get("Activity", "") or "").strip() if pd.notna(row.get("Activity")) else "",
            "name": str(row.get("Name", "") or "").strip() if pd.notna(row.get("Name")) else "",
            "sex": str(row.get("Sex", "") or "").strip() if pd.notna(row.get("Sex")) else "",
            "age": str(row.get("Age", "") or "").strip() if pd.notna(row.get("Age")) else "",
            "injury": str(row.get("Injury", "") or "").strip() if pd.notna(row.get("Injury")) else "",
            "fatal": fatal_raw.startswith("Y"),
            "time": str(row.get("Time", "") or "").strip() if pd.notna(row.get("Time")) else "",
            "species": species[:90],
            "source": "GSAF",
            "source_detail": "Global Shark Attack File (Shark Research Institute)",
            "lat": round(lat, 5),
            "lon": round(lon, 5),
        })
    return records


def process_california(zip_bytes: bytes) -> list[dict]:
    print("Processing California CDFW …")
    records = []
    with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
        xlsx_name = [n for n in zf.namelist() if n.endswith(".xlsx")][0]
        with zf.open(xlsx_name) as f:
            df = pd.read_excel(f)
    print(f"  CA raw rows: {len(df)}")

    for _, row in df.iterrows():
        try:
            lat = float(row.get("Latitude")) if pd.notna(row.get("Latitude")) else None
            lon = float(row.get("Longitude")) if pd.notna(row.get("Longitude")) else None
        except (TypeError, ValueError):
            lat = lon = None
        if lat is None or lon is None:
            continue
        # skip bad coords
        if not (30 < lat < 43 and -130 < lon < -114):
            continue

        dt = row.get("Date")
        year = None
        date_str = ""
        if pd.notna(dt):
            try:
                if hasattr(dt, "year"):
                    year = int(dt.year)
                    date_str = dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10]
                else:
                    date_str = str(dt)
                    year = int(str(dt)[:4])
            except Exception:
                pass

        injury = str(row.get("Injury", "") or "").lower()
        fatal = "fatal" in injury and "non" not in injury

        species = str(row.get("Species", "") or "").strip()
        if species.lower() in ("nan", "none", ""):
            species = ""

        records.append({
            "id": f"cdfw-{row.get('IncidentNum', len(records)+1)}",
            "date": date_str,
            "year": year or 0,
            "type": "Unprovoked",  # CDFW does not always classify; default
            "country": "USA",
            "state": "California",
            "location": str(row.get("Location", "") or "").strip(),
            "activity": str(row.get("Mode", "") or "").strip(),
            "name": "",
            "sex": "",
            "age": "",
            "injury": str(row.get("Injury", "") or "").strip(),
            "fatal": fatal,
            "time": str(row.get("Time", "") or "").strip(),
            "species": species[:90],
            "source": "CDFW California",
            "source_detail": "California Department of Fish & Wildlife Shark Incident Database",
            "lat": round(lat, 5),
            "lon": round(lon, 5),
        })
    print(f"  CA records with valid coords: {len(records)}")
    return records


def main() -> int:
    print("=" * 60)
    print("JAWS STIMULATOR — multi-source data update")
    print(f"UTC: {datetime.now(timezone.utc).isoformat()}")
    print("=" * 60)

    try:
        print("Downloading GSAF …")
        gsaf_bytes = download(GSAF_URL)
        gsaf = process_gsaf(gsaf_bytes)

        print("Downloading California CDFW …")
        try:
            ca_bytes = download(CA_ZIP_URL)
            ca = process_california(ca_bytes)
        except Exception as e:
            print(f"  CA download/process failed (continuing without): {e}")
            ca = []

        # Merge: prefer GSAF as canonical; CA adds precise CA points
        # Tag duplicates loosely by year+location for transparency but keep both
        # so users can filter by source. Assign sequential display ids.
        all_recs = gsaf + ca
        for i, r in enumerate(all_recs, 1):
            r["id"] = i

        OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(all_recs, f, ensure_ascii=False)

        years = [r["year"] for r in all_recs if r.get("year")]
        sources = {}
        for r in all_recs:
            sources[r["source"]] = sources.get(r["source"], 0) + 1

        meta = {
            "updated": datetime.now(timezone.utc).isoformat(),
            "sources": sources,
            "count": len(all_recs),
            "min_year": min(years) if years else MIN_YEAR,
            "max_year": max(years) if years else None,
            "note": "GSAF = Global Shark Attack File. CDFW = California Dept of Fish & Wildlife (real coordinates).",
            "reference": {
                "shark_data_lab_total": 7207,
                "shark_data_lab_fatal": 1523,
                "shark_data_lab_fatality_rate": "21.1%",
                "shark_data_lab_countries": 121,
            },
        }
        with open(META_JSON, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)

        print(f"Wrote {len(all_recs)} records → {OUT_JSON}")
        print(f"Sources: {sources}")
        print(f"Years: {meta['min_year']}–{meta['max_year']}")
        print("Done.")
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
