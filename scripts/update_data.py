#!/usr/bin/env python3
"""
JAWS STIMULATOR — Daily GSAF data updater
Downloads the latest Global Shark Attack File (Excel), cleans it,
adds approximate coastal coordinates, and writes data/shark_data.json.
"""

import json
import os
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
GSAF_URL = "https://www.sharkattackfile.net/spreadsheets/GSAF5.xls"
ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "data" / "shark_data.json"
TMP_XLS = ROOT / "data" / "_gsaf_tmp.xls"
MAX_RECORDS = 2500          # keep browser responsive
MIN_YEAR = 1990             # focus on modern era + recent updates
RANDOM_SEED = 42

# Approximate coastal / high-activity points (Shark Data Lab does proper geocoding)
CENTROIDS = {
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
    "HAWAII": (20.5, -156.5),  # treated as USA state usually
}


def download_gsaf(url: str, dest: Path) -> None:
    print(f"Downloading GSAF from {url} …")
    r = requests.get(url, timeout=120, headers={"User-Agent": "JawsStimulator/1.0 (GitHub Action)"})
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


def get_coords(country: str) -> tuple[float, float]:
    c = normalize_country(country)
    if c in CENTROIDS:
        return CENTROIDS[c]
    # partial match
    for k, v in CENTROIDS.items():
        if k in c or c in k:
            return v
    # last resort: tropical ocean
    return (0.0, 0.0)


def process(xls_path: Path) -> list[dict]:
    print("Reading Excel …")
    try:
        df = pd.read_excel(xls_path, engine="xlrd")
    except Exception:
        df = pd.read_excel(xls_path, engine="openpyxl")

    df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(df.columns)]
    print(f"  Columns: {list(df.columns)[:12]}…")
    print(f"  Raw rows: {len(df)}")

    # Year cleaning
    if "Year" not in df.columns:
        raise RuntimeError("No 'Year' column found in GSAF file")
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df = df.dropna(subset=["Year"])
    df = df[df["Year"] >= MIN_YEAR]
    df = df.sort_values("Year", ascending=False)

    # Keep most recent MAX_RECORDS
    if len(df) > MAX_RECORDS:
        df = df.head(MAX_RECORDS)
    print(f"  After year filter & limit: {len(df)} rows")

    random.seed(RANDOM_SEED)
    records = []
    for idx, row in df.iterrows():
        country_raw = str(row.get("Country", "") or "")
        country = normalize_country(country_raw)
        base_lat, base_lon = get_coords(country)

        # Small random offset so points don't stack perfectly
        lat = base_lat + random.uniform(-3.5, 3.5)
        lon = base_lon + random.uniform(-5.0, 5.0)
        lat = max(-85.0, min(85.0, lat))
        lon = max(-180.0, min(180.0, lon))

        fatal_raw = str(row.get("Fatal Y/N", "") or "").upper().strip()
        fatal = fatal_raw.startswith("Y")

        species = str(row.get("Species", "") or row.get("Species ", "") or "").strip()
        if species.lower() in ("nan", "none", ""):
            species = ""

        rec = {
            "id": len(records) + 1,
            "date": str(row.get("Date", "") or "").strip(),
            "year": int(row["Year"]),
            "type": str(row.get("Type", "") or "").strip(),
            "country": country_raw.strip() or country,
            "state": str(row.get("State", "") or "").strip() if pd.notna(row.get("State")) else "",
            "location": str(row.get("Location", "") or "").strip() if pd.notna(row.get("Location")) else "",
            "activity": str(row.get("Activity", "") or "").strip() if pd.notna(row.get("Activity")) else "",
            "name": str(row.get("Name", "") or "").strip() if pd.notna(row.get("Name")) else "",
            "sex": str(row.get("Sex", "") or "").strip() if pd.notna(row.get("Sex")) else "",
            "age": str(row.get("Age", "") or "").strip() if pd.notna(row.get("Age")) else "",
            "injury": str(row.get("Injury", "") or "").strip() if pd.notna(row.get("Injury")) else "",
            "fatal": fatal,
            "time": str(row.get("Time", "") or "").strip() if pd.notna(row.get("Time")) else "",
            "species": species[:80],
            "source": str(row.get("Source", "") or "").strip()[:140] if pd.notna(row.get("Source")) else "",
            "lat": round(lat, 4),
            "lon": round(lon, 4),
        }
        records.append(rec)

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
        payload = {
            "updated": datetime.now(timezone.utc).isoformat(),
            "source": "Global Shark Attack File (GSAF) — https://www.sharkattackfile.net",
            "note": "Coordinates are approximate coastal points for visualization. Shark Data Lab performs full geocoding.",
            "count": len(records),
            "records": records,
        }

        # Keep backward-compatible flat array for the current frontend
        # while also writing metadata. Frontend currently expects a list.
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=None)

        # Also write a meta file the dashboard can optionally show
        meta_path = OUT_JSON.with_name("meta.json")
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump({
                "updated": payload["updated"],
                "source": payload["source"],
                "count": payload["count"],
                "min_year": MIN_YEAR,
                "max_records": MAX_RECORDS,
            }, f, indent=2)

        print(f"Wrote {len(records)} records → {OUT_JSON}")
        print(f"Meta → {meta_path}")

        # Cleanup temp
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
