# JAWS STIMULATOR

**Global Shark Encounter Dashboard** — a sleek, Martian-inspired data screen for visualizing recorded shark–human incidents worldwide.

![Theme](https://img.shields.io/badge/theme-The%20Martian-00e5ff?style=flat-square)
![Data](https://img.shields.io/badge/data-GSAF-ff2d95?style=flat-square)
![Update](https://img.shields.io/badge/update-daily-brightgreen?style=flat-square)

Live demo (after you enable GitHub Pages):  
`https://<your-username>.github.io/jaws-stimulator/`

---

## Features

- Dark, cyan-accented interface inspired by the data screens in *The Martian*
- Interactive world map (Leaflet + Carto dark tiles)
- Click any marker or list row → floating detail panel
- Full-text search (location, species, activity, country, name…)
- Filters: All / Fatal / Non-fatal / Unprovoked / 2024+
- Reset view + zoom controls
- Year trend, outcome distribution, top countries & species
- **Automatic daily data refresh** via GitHub Actions

## Data source

All incident records come from the **Global Shark Attack File (GSAF)** maintained by the Shark Research Institute:

- https://www.sharkattackfile.net  
- https://www.sharkdatalab.com/en (cleaned / geocoded presentation of the same underlying data)

Coordinates in this dashboard are approximate coastal points used for visualization. Shark Data Lab performs proper geocoding on the full dataset.

## Deploy on GitHub Pages (5 minutes)

1. Create a new repository on GitHub named e.g. `jaws-stimulator`.
2. Upload / push the contents of this folder (do **not** nest an extra folder).
3. Go to **Settings → Pages**.
4. Under **Source** choose **Deploy from a branch**.
5. Select branch `main` (or `master`) and folder `/ (root)`.
6. Save. After a minute your site will be live at  
   `https://<username>.github.io/jaws-stimulator/`

### Enable the daily auto-update

The workflow file `.github/workflows/update-shark-data.yml` is already included.

1. After pushing, go to the **Actions** tab of your repository.
2. If GitHub asks you to enable workflows, click **I understand…** / enable them.
3. You can also click **Daily GSAF Data Update → Run workflow** to trigger an immediate refresh.

The action runs every day at 06:15 UTC, downloads the latest GSAF Excel, regenerates `data/shark_data.json` + `data/meta.json`, and commits the result if anything changed.

> **Note:** The first time the action runs it needs write permission (already granted via `permissions: contents: write`). If your organization has restricted Actions, you may need to allow the workflow manually.

## Local development

```bash
# Optional: refresh data yourself
pip install pandas openpyxl xlrd requests
python scripts/update_data.py

# Serve locally
python -m http.server 8080
# open http://localhost:8080
```

## File structure

```
.
├── index.html                 # Dashboard UI
├── data/
│   ├── shark_data.json        # Incident records (auto-updated)
│   └── meta.json              # Last-update metadata
├── scripts/
│   └── update_data.py         # GSAF downloader + cleaner
└── .github/workflows/
    └── update-shark-data.yml  # Daily scheduled job
```

## Credits

- Incident data © Shark Research Institute / Global Shark Attack File contributors  
- Visual language inspired by the film *The Martian* (2015)  
- Map tiles © CARTO / OpenStreetMap contributors  

---

**JAWS STIMULATOR** — stay curious, stay informed.
