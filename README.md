# PCT Waiting Time Dashboard

A self-contained dashboard visualising Plant Cycle Time (PCT) waiting times across Opella's manufacturing and packaging operations.

**Project 1000 target: 17 days PCT** (current avg ~34 days)

## Live dashboard

Deployed via GitHub Pages → [michelleko22.github.io/PCT-Dashboard](https://michelleko22.github.io/PCT-Dashboard)

## What it shows

- **PCT Overview** — KPI cards, daily batch cycle time breakdown by product type (Coated Tablet / Uncoated Tablet / Hard Capsule / Softgel), top bottlenecks, and batch-level detail table
- **Work Centre Detail** — Compression booth status snapshot (8 booths), idle time, current and recent jobs

## Regenerating the dashboard

The dashboard is a single static HTML file (`index.html`) generated from the production schedule Excel files.

### Requirements

```
pip install openpyxl
```

### Setup

Place the source Excel files in the `data/` folder (not committed — see note below):

```
data/
├── PRODUCTION SCHED - Manufacturing.xlsm
└── PRODUCTION SCHED - Packaging.xlsm
```

### Run

```bash
python3 generate_dashboard.py
```

This overwrites `index.html` with fresh data. Commit and push to update the live site.

## Repository structure

```
PCT-Dashboard/
├── index.html                  # Generated dashboard (GitHub Pages entry point)
├── generate_dashboard.py       # Data extraction + HTML generation script
├── docs/
│   ├── design-mockup.docx      # Original design spec
│   └── mockups/                # Reference screenshots
├── data/                       # ← gitignored; place Excel files here locally
└── .gitignore
```

> **Note:** The `data/` folder is excluded from version control. The Excel production schedules contain operational data and are 23 MB combined.
