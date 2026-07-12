<img src="assets/logo.png" alt="The 529 Network" height="56">

# Prepaid Committee Data Tool

Internal Streamlit app for The 529 Network's Prepaid Committee. It holds the consolidated
record of prepaid tuition plan statistics from 2005 to the present, lets the committee
explore and QA every plan-year, exports committee-format workbooks, and accepts each new
annual collection so the record keeps growing.

**22 plan entities · 18 reporting periods · 350 rows**

---

## Project structure

```
prepaid_committee_tool/
├── streamlit_app.py                  the app
├── requirements.txt
├── README.md
├── .streamlit/config.toml            brand theme
├── src/
│   └── prepaid_parser.py             header-driven parser for committee Excel files
├── data/                             source of truth
│   ├── prepaid_master.csv            one row per plan per reporting period
│   ├── prepaid_meta.json             plan registry + static attributes
│   └── prepaid_attrs_by_year.json    per-year descriptive + actuarial attributes
├── assets/
│   ├── logo.png                      full logo (transparent)
│   ├── logo_small.png                app header
│   ├── logo_excel.png                embedded in Excel exports
│   └── favicon.png                   browser tab icon
└── samples/                          example exports
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy (Streamlit Community Cloud)

1. Push this folder to a GitHub repo.
2. On share.streamlit.io, create an app pointing at `streamlit_app.py`.
3. No secrets or configuration needed. Internal use, open link.

---

## The tabs

| Tab | What it does |
|---|---|
| **Overview** | Headline totals and a plan-by-year coverage heatmap. Green reported, amber carried-forward or lagged as-of, blank not collected. |
| **Plan Profile** | One plan: attribute cards, funded / assets / accounts history, and a full table with as-of dates and notes. |
| **Compare** | Sortable multi-plan comparison on any metric, any year. |
| **Trends** | Any metric across years, several plans overlaid. |
| **Committee Era (2022+)** | The years carrying the full detail set. Single-year detail, actuarial assumptions across years, and a metric matrix. |
| **Data Quality** | Completeness by metric, flags to review, and interior missing plan-years. |
| **Upload & Append** | Add a new year from a committee Excel file. |
| **Exports** | Committee-format workbooks, the analytical master, and filtered CSV. |

## Adding a new year (or the missing 2016 / 2017 / 2019)

1. Open **Upload & Append**, set the reporting year, and upload the committee Excel file.
   Both the older (2022–23) and newer (2024–25) column layouts are handled; Michigan
   MET I/II and Mississippi tiers split automatically. A blank template is downloadable
   there if you want submissions in a fixed shape.
2. Review the parsed preview. Choose overwrite or skip if the year already exists, then merge.
3. Download **Master (CSV)** and commit it to `data/prepaid_master.csv`.

That commit is what makes the year permanent. Streamlit Community Cloud does not persist
in-app writes, so the file in `data/` is the source of truth.

## Exports

- **Committee format** — reproduces the committee sheet layout (merged headers, OPEN /
  CLOSED sections, Michigan and Mississippi written in the committee's own cell style).
  Pick one year for a single sheet, or several to get one workbook with a tab per year.
- **Analytical master** — README, the long table, and Funded / Assets / Accounts matrices
  across every year.
- **Filtered CSV** — any slice of years, statuses, and plans.

All Excel exports carry the brand logo.

---

## Data conventions

- Funded status is stored as a fraction and displayed as a percent (`1.57` = 157%).
- Assets and payouts are in millions of dollars. Accounts are counts.
- Michigan MET is two entities: MET I (Plans B & C, closed) and MET II (Plan D, open).
- Mississippi MPACT reports one account base with funded status split by tier; the Horizon
  figure is the headline value and Legacy is captured in notes.
- Texas runs on an August 31 fiscal year. Texas has two plans: the Guaranteed Tuition Plan
  (closed) and the Tuition Promise Fund (open).
- Where quarterly workbooks overlapped, the most recent file's value was used, since later
  files carry restated actuarial figures.

## Known gaps

- **2016, 2017, and 2019** were never collected and are absent from the record.
- Descriptive and actuarial columns (tuition growth, investment return, benefit structure)
  exist only from **2022** on. Earlier years carry funded status, assets, and accounts.
  Accounts-since-inception and payout figures begin with the 2018 one-pagers.
- Maryland reported no data for 2024 or 2025.
