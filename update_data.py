#!/usr/bin/env python3
"""
JAWS STIMULATOR — Multi-source shark incident data updater

GSAF bulk Excel (GSAF5.xls) was removed from public hosting (404 as of 2026-10).
This script:
  1. Tries several known GSAF / mirror URLs
  2. Falls back to the existing data/shark_data.json if download fails
  3. Always attempts California CDFW refresh (real lat/lon)
  4. Merges Wikipedia US-fatal list if present
  5. Exits 0 when usable data remains (so GitHub Actions does not fail red)
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

# Candidate GSAF bulk file URLs (first success wins)
GSAF_CANDIDATE_URLS = [
    "https://www.sharkattackfile.net/spreadsheets/GSAF5.xls",
    "https://sharkattackfile.net/spreadsheets/GSAF5.xls",
    "https://www.sharkattackfile.net/spreadsheets/GSAF5.xlsx",
    # Historical community mirrors (may lag official)
    "https://raw.githubusercontent.com/thedevastator/shark-attacks-the-risks-of-coastal-water-activit/main/GSAF5.xls.csv",
]

CA_ZIP_URL = (
    "https://data.cnra.ca.gov/dataset/022f670f-eb4e-4314-9fad-74f7f9560400/"
    "resource/fa77215c-a06c-4fa2-b8bc-ffbcaaa65286/download/sharkincidents_1950_2022_220302.zip"
)

ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "data" / "shark_data.json"
META_JSON = ROOT / "data" / "meta.json"
WIKI_JSON = ROOT / "data" / "wiki_us_fatal.json"
MAX_GSAF = 8000
MIN_YEAR = 1800
RANDOM_SEED = 42
UA = "JawsStimulator/2.1 (GitHub Action; educational dashboard; contact via repo issues)"

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

CONTINENT = {
    "USA": "North America", "UNITED STATES": "North America", "MEXICO": "North America",
    "CANADA": "North America", "BAHAMAS": "North America", "CUBA": "North America",
    "JAMAICA": "North America", "COSTA RICA": "North America", "PANAMA": "North America",
    "BERMUDA": "North America", "PUERTO RICO": "North America",
    "BRAZIL": "South America", "ECUADOR": "South America", "COLOMBIA": "South America",
    "VENEZUELA": "South America", "CHILE": "South America", "ARGENTINA": "South America", "PERU": "South America",
    "AUSTRALIA": "Oceania", "NEW ZEALAND": "Oceania", "FIJI": "Oceania", "PAPUA NEW GUINEA": "Oceania",
    "NEW CALEDONIA": "Oceania", "FRENCH POLYNESIA": "Oceania", "SOLOMON ISLANDS": "Oceania",
    "TONGA": "Oceania", "VANUATU": "Oceania", "SAMOA": "Oceania", "GUAM": "Oceania",
    "SOUTH AFRICA": "Africa", "EGYPT": "Africa", "MOZAMBIQUE": "Africa", "REUNION": "Africa",
    "SENEGAL": "Africa", "KENYA": "Africa", "MAURITIUS": "Africa", "SEYCHELLES": "Africa",
    "ITALY": "Europe", "FRANCE": "Europe", "SPAIN": "Europe", "UNITED KINGDOM": "Europe", "TURKEY": "Europe",
    "JAPAN": "Asia", "INDIA": "Asia", "PHILIPPINES": "Asia", "INDONESIA": "Asia",
    "THAILAND": "Asia", "VIETNAM": "Asia", "SRI LANKA": "Asia", "IRAN": "Asia", "HONG KONG": "Asia",
    "MALDIVES": "Asia",
}


def continent_for(country: str) -> str:
    c = (country or "").upper().strip()
    if c in CONTINENT:
        return CONTINENT[c]
    for k, v in CONTINENT.items():
        if k in c or c in k:
            return v
    return "Unknown"


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


def download_bytes(url: str, timeout: int = 90) -> bytes:
    r = requests.get(url, timeout=timeout, headers={"User-Agent": UA}, allow_redirects=True)
    r.raise_for_status()
    if len(r.content) < 1000:
        raise ValueError(f"Response too small ({len(r.content)} bytes) from {url}")
    return r.content


def try_download_gsaf() -> tuple[bytes | None, str | None]:
    last_err = None
    for url in GSAF_CANDIDATE_URLS:
        try:
            print(f"  Trying {url} …")
            data = download_bytes(url)
            print(f"  OK — {len(data):,} bytes from {url}")
            return data, url
        except Exception as e:
            print(f"  Failed: {e}")
            last_err = e
    print(f"  All GSAF candidates failed (last: {last_err})")
    return None, None


def read_gsaf_dataframe(raw: bytes, source_url: str) -> pd.DataFrame:
    """Read Excel or CSV GSAF payload."""
    if source_url.endswith(".csv") or raw[:200].lstrip().startswith(b"Case") or b"," in raw[:200]:
        try:
            return pd.read_csv(BytesIO(raw), low_memory=False)
        except Exception:
            pass
    for engine in ("xlrd", "openpyxl"):
        try:
            return pd.read_excel(BytesIO(raw), engine=engine)
        except Exception:
            continue
    # last resort: csv
    return pd.read_csv(BytesIO(raw), low_memory=False)


def process_gsaf(raw: bytes, source_url: str) -> list[dict]:
    print("Processing GSAF …")
    df = read_gsaf_dataframe(raw, source_url)
    df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(df.columns)]
    print(f"  Columns: {list(df.columns)[:12]}…")
    print(f"  Raw rows: {len(df)}")

    year_col = "Year" if "Year" in df.columns else None
    if year_col is None:
        for c in df.columns:
            if str(c).lower() == "year":
                year_col = c
                break
    if year_col is None:
        raise RuntimeError("No Year column in GSAF file")

    df[year_col] = pd.to_numeric(df[year_col], errors="coerce")
    df = df.dropna(subset=[year_col])
    df = df[(df[year_col] >= MIN_YEAR) & (df[year_col] <= datetime.now().year + 1)]
    df = df.sort_values(year_col, ascending=False)
    if len(df) > MAX_GSAF:
        df = df.head(MAX_GSAF)
    print(f"  Kept rows: {len(df)}")

    def col(*names):
        for n in names:
            if n in df.columns:
                return n
        for n in names:
            for c in df.columns:
                if str(c).strip().lower() == n.lower():
                    return c
        return None

    c_date = col("Date")
    c_type = col("Type")
    c_country = col("Country")
    c_state = col("State", "Area")
    c_loc = col("Location")
    c_act = col("Activity")
    c_name = col("Name")
    c_sex = col("Sex", "Sex ")
    c_age = col("Age")
    c_injury = col("Injury")
    c_fatal = col("Fatal Y/N", "Fatal (Y/N)", "Fatal")
    c_time = col("Time")
    c_species = col("Species", "Species ")
    c_source = col("Source")

    random.seed(RANDOM_SEED)
    records = []
    for _, row in df.iterrows():
        country_raw = str(row.get(c_country, "") or "").strip() if c_country else ""
        state = str(row.get(c_state, "") or "").strip() if c_state and pd.notna(row.get(c_state)) else ""
        location = str(row.get(c_loc, "") or "").strip() if c_loc and pd.notna(row.get(c_loc)) else ""
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

        fatal_raw = str(row.get(c_fatal, "") or "").upper().strip() if c_fatal else ""
        fatal = fatal_raw.startswith("Y")

        species = str(row.get(c_species, "") or "").strip() if c_species and pd.notna(row.get(c_species)) else ""
        if species.lower() in ("nan", "none"):
            species = ""

        attack_type = str(row.get(c_type, "") or "").strip() if c_type and pd.notna(row.get(c_type)) else ""
        if attack_type.lower() in ("nan", "none"):
            attack_type = ""

        records.append({
            "id": 0,  # assigned later
            "date": str(row.get(c_date, "") or "").strip() if c_date else "",
            "year": int(row[year_col]),
            "type": attack_type,
            "country": country_raw or normalize_country(country_raw),
            "state": state,
            "location": location,
            "activity": str(row.get(c_act, "") or "").strip() if c_act and pd.notna(row.get(c_act)) else "",
            "name": str(row.get(c_name, "") or "").strip() if c_name and pd.notna(row.get(c_name)) else "",
            "sex": str(row.get(c_sex, "") or "").strip() if c_sex and pd.notna(row.get(c_sex)) else "",
            "age": str(row.get(c_age, "") or "").strip() if c_age and pd.notna(row.get(c_age)) else "",
            "injury": str(row.get(c_injury, "") or "").strip() if c_injury and pd.notna(row.get(c_injury)) else "",
            "fatal": fatal,
            "time": str(row.get(c_time, "") or "").strip() if c_time and pd.notna(row.get(c_time)) else "",
            "species": species[:90],
            "source": "GSAF",
            "source_detail": "Global Shark Attack File (Shark Research Institute)",
            "continent": continent_for(country_raw),
            "lat": round(lat, 5),
            "lon": round(lon, 5),
        })
    return records


def process_california(zip_bytes: bytes) -> list[dict]:
    print("Processing California CDFW …")
    records = []
    with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
        xlsx_name = [n for n in zf.namelist() if n.lower().endswith((".xlsx", ".xls"))][0]
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
            "id": 0,
            "date": date_str,
            "year": year or 0,
            "type": "Unprovoked",
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
            "continent": "North America",
            "lat": round(lat, 5),
            "lon": round(lon, 5),
        })
    print(f"  CA records with valid coords: {len(records)}")
    return records


def load_existing_gsaf_records() -> list[dict]:
    """Keep prior GSAF rows when live download is unavailable."""
    if not OUT_JSON.exists():
        return []
    try:
        data = json.loads(OUT_JSON.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        kept = [r for r in data if r.get("source") == "GSAF"]
        print(f"  Loaded {len(kept)} existing GSAF records from {OUT_JSON}")
        return kept
    except Exception as e:
        print(f"  Could not load existing data: {e}")
        return []


def load_wiki() -> list[dict]:
    if not WIKI_JSON.exists():
        return []
    try:
        wiki = json.loads(WIKI_JSON.read_text(encoding="utf-8"))
        for r in wiki:
            r.setdefault("continent", "North America")
            r.setdefault("source", "Wikipedia (US Fatal)")
        print(f"  Loaded {len(wiki)} Wikipedia US-fatal records")
        return wiki
    except Exception as e:
        print(f"  Wiki load failed: {e}")
        return []


def main() -> int:
    print("=" * 60)
    print("JAWS STIMULATOR — multi-source data update")
    print(f"UTC: {datetime.now(timezone.utc).isoformat()}")
    print("=" * 60)

    gsaf_records: list[dict] = []
    gsaf_source = "existing-cache"
    gsaf_bytes, gsaf_url = try_download_gsaf()
    if gsaf_bytes:
        try:
            gsaf_records = process_gsaf(gsaf_bytes, gsaf_url or "")
            gsaf_source = gsaf_url or "download"
        except Exception as e:
            print(f"  GSAF parse failed: {e}")
            gsaf_records = load_existing_gsaf_records()
            gsaf_source = "existing-cache (parse-failed)"
    else:
        print("GSAF public bulk file unavailable — using cached GSAF rows.")
        gsaf_records = load_existing_gsaf_records()
        gsaf_source = "existing-cache"

    ca_records: list[dict] = []
    try:
        print("Downloading California CDFW …")
        ca_bytes = download_bytes(CA_ZIP_URL)
        ca_records = process_california(ca_bytes)
    except Exception as e:
        print(f"  CA download/process failed (continuing): {e}")
        # keep prior CA rows if any
        if OUT_JSON.exists():
            try:
                prev = json.loads(OUT_JSON.read_text(encoding="utf-8"))
                ca_records = [r for r in prev if r.get("source") == "CDFW California"]
                print(f"  Kept {len(ca_records)} prior CA rows")
            except Exception:
                pass

    wiki_records = load_wiki()

    all_recs = gsaf_records + ca_records + wiki_records
    for i, r in enumerate(all_recs, 1):
        r["id"] = i
        r.setdefault("continent", continent_for(r.get("country", "")))

    if not all_recs:
        print("ERROR: No records available from any source.", file=sys.stderr)
        return 1

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(all_recs, f, ensure_ascii=False)

    years = [r["year"] for r in all_recs if r.get("year")]
    sources: dict[str, int] = {}
    for r in all_recs:
        sources[r.get("source", "?")] = sources.get(r.get("source", "?"), 0) + 1

    meta = {
        "updated": datetime.now(timezone.utc).isoformat(),
        "sources": sources,
        "count": len(all_recs),
        "min_year": min(years) if years else MIN_YEAR,
        "max_year": max(years) if years else None,
        "gsaf_fetch": gsaf_source,
        "note": (
            "GSAF public bulk Excel (GSAF5.xls) returned 404 from sharkattackfile.net as of 2026-10. "
            "Updater tries multiple URLs then falls back to the last successful GSAF snapshot in-repo. "
            "CDFW California supplies precise CA coordinates. Wikipedia covers historical US fatal cases."
        ),
        "reference": {
            "shark_data_lab_total": 6528,
            "shark_data_lab_fatal": 1480,
            "shark_data_lab_countries": 119,
            "isaf_note": "Florida Museum ISAF individual case files are not a public bulk download.",
        },
    }
    with open(META_JSON, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

    print(f"Wrote {len(all_recs)} records → {OUT_JSON}")
    print(f"Sources: {sources}")
    print(f"GSAF fetch mode: {gsaf_source}")
    print(f"Years: {meta['min_year']}–{meta['max_year']}")
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
