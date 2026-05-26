---
name: cv-paper-pipeline
description: >
  Search multiple academic sources (CVF Open Access, arXiv, IEEE Xplore, Web of
  Science, ScienceDirect, Scopus) for Computer Vision papers in Anomaly
  Detection / Object Detection / Image Classification, collect them into the
  project's Excel workbooks, then regenerate the literature-viewer website
  (index.html) from those workbooks. The pipeline runs automatically every
  Wednesday and Friday at 02:00 via Windows Task Scheduler and emails a summary
  to azaz31855@gmail.com AND fctien@ntut.edu.tw after each run (multi-recipient
  via comma-separated NOTIFY_TO in scripts/.env). Use this skill when the user
  wants to update the paper database, add papers from a new source, pull the
  latest conference proceedings, rebuild the website data, or change the
  schedule / email settings.
---

# CV Paper Pipeline

End-to-end workflow: **search sources → Excel → website**.

```
CVF / arXiv / IEEE / WoS / ScienceDirect / Scopus
        │  (programmatic where possible, manual export otherwise)
        ▼
   Excel workbooks (.xlsx)  ← single source of truth
        │  regenerate_*_json()
        ▼
   *_data.json  →  injected into index.html
        │  git push
        ▼
   GitHub Pages (auto-deploy)
```

## Project layout

| File | Role |
| --- | --- |
| `AnomalyDetection_Papers_Summary_v10_20260425.xlsx` | AD source workbook |
| `Object_Detection_Papers_Ranking_2021_2026.xlsx`     | OD source workbook |
| `Image_Classification_Papers_Ranking_2021_2026.xlsx` | CLS source workbook |
| `Anomaly_Synthesis_Papers_Benchmark_2021_2026.xlsx`  | AS source workbook (domain disabled on site) |
| `cls_supplemental.json` | Hand-curated CLS rows merged on top of the xlsx |
| `index.html` | Single-file website; data is injected as `const *_DATA = …` |
| `scripts/update_papers.py` | Main pipeline — arXiv + CVF fetch, regenerate, inject, push, email. Run by Task Scheduler Wed/Fri 02:00. |
| `scripts/run_update.bat` | Task Scheduler wrapper |
| `scripts/.env` | SMTP credentials (gitignored) |
| `.claude/skills/cv-paper-pipeline/scripts/fetch_cvf.py` | Standalone CVF scraper (for ad-hoc / specific-conference scans) |

## Step 1 — Search the sources

### 1a. Programmatic sources (run directly)

**arXiv** — already wired into `scripts/update_papers.py`. To pull the latest:
```bash
python scripts\update_papers.py --dry        # preview only
python scripts\update_papers.py               # full run + push + email
```

**CVF Open Access** — public HTML. The scheduled pipeline scans it
automatically (`process_cvf()` in `update_papers.py`, see *Automation*
below). For an **ad-hoc scan of a specific conference**, use the standalone
script:
```bash
python .claude\skills\cv-paper-pipeline\scripts\fetch_cvf.py --conf CVPR --year 2025
python .claude\skills\cv-paper-pipeline\scripts\fetch_cvf.py --conf ICCV --year 2025 --domain ad
```
Produces `cvf_results.csv` with `domain, conference, year, title, authors,
paper_url, dataset_guess, …`. Review it, then merge into the Excel workbooks
(Step 2).

### 1b. Login-gated sources (manual export)

IEEE Xplore, Web of Science, ScienceDirect and Scopus **cannot be scraped** —
they need an institutional login or a paid API key. The supported path is a
**manual metadata export** while on the school VPN:

1. Open `references/source_access.md` for the exact search strings and the
   click-path to export a CSV/BibTeX from each site.
2. Save each export under `sources/` (e.g. `sources/ieee_export.csv`).
3. Tell this skill the file path; it will parse the CSV and merge the rows.

Never bulk-download PDFs, never bypass CAPTCHAs or logins. Abstract reading +
metadata export only.

## Step 2 — Merge into Excel

The `.xlsx` files are the **single source of truth**. For each new paper:

1. Pick the workbook by domain (AD / OD / CLS).
2. Append a row to the relevant per-dataset sheet **and** the "all papers"
   sheet, following the exact column order in `references/excel_schema.md`.
3. **Metrics policy** — only fill a metric number if it is verified from the
   paper's own results table. If unverified, leave it blank and write the
   reason in the notes column. Never fabricate numbers (see the OSD-IRF
   incident — wrong numbers were caught by the user).
4. If a workbook is open in Excel it will be **locked**; either ask the user to
   close it, or stage the rows in a `*_supplemental.json` file (see
   `cls_supplemental.json` for the pattern) which `regenerate_*_json()` merges
   on top of the xlsx.

## Step 3 — Regenerate the website

```bash
python -c "import sys; sys.path.insert(0,'scripts'); import update_papers as up; \
ad=up.regenerate_ad_json(); od=up.regenerate_od_json(); \
cls=up.regenerate_cls_json(); asd=up.regenerate_as_json(); \
up.reinject_html(ad,od,cls,asd); print('done')"
```

This reads every workbook, rebuilds `*_data.json`, and rewrites the
`const *_DATA = …` blocks inside `index.html`. Then commit + push:

```bash
git add index.html *_data.json *.xlsx cls_supplemental.json
git commit -m "update: <what changed>"
git push
```

GitHub Pages redeploys automatically.

## Automation — scheduled runs

The whole pipeline runs **unattended every Wednesday and Friday at 02:00**
via Windows Task Scheduler, and **emails a summary to two recipients
(azaz31855@gmail.com + fctien@ntut.edu.tw)** after each run.

### What a scheduled run does
`scripts/update_papers.py` (invoked by `scripts/run_update.bat`) executes:
1. **arXiv** fetch — last 7 days, AD / OD / CLS / AS queries.
2. **CVF Open Access** scan — `process_cvf()` scrapes CVPR / ICCV / ECCV /
   WACV proceedings, dedups by method name, appends genuinely-new papers to
   the "All Papers" sheets (metrics left blank — pending verification).
3. Append rows to the `.xlsx` workbooks.
4. Regenerate `*_data.json` and re-inject into `index.html`.
5. `git push` → GitHub Pages redeploys.
6. **Email** a summary (new papers per domain, arXiv + CVF) to **both**
   `azaz31855@gmail.com` and `fctien@ntut.edu.tw` in a single send
   (multi-recipient via comma-separated `NOTIFY_TO`).

### Task Scheduler entries
Two weekly tasks (run as the logged-in user):

| Task name | Trigger |
| --- | --- |
| `AD-OD Paper Auto-Update Wed` | Weekly · Wednesday · 02:00 |
| `AD-OD Paper Auto-Update Fri` | Weekly · Friday · 02:00 |

Register / re-register them (run in an elevated shell):
```
schtasks /Create /TN "AD-OD Paper Auto-Update Wed" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D WED /ST 02:00 /F
schtasks /Create /TN "AD-OD Paper Auto-Update Fri" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D FRI /ST 02:00 /F
```
Verify: `schtasks /Query /TN "AD-OD Paper Auto-Update Wed"`
Run on demand: `schtasks /Run /TN "AD-OD Paper Auto-Update Wed"`
Change the time: re-create with a different `/ST`. Change the days: edit `/D`.

### Email configuration
SMTP credentials live in `scripts/.env` (gitignored):
```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=azaz31855@gmail.com
SMTP_PASSWORD=<gmail app password>
NOTIFY_TO=azaz31855@gmail.com,fctien@ntut.edu.tw
```
`NOTIFY_TO` is **comma- or whitespace-separated** — list as many recipients as
needed. The script parses it via regex split (`re.split(r'[,;\s]+', ...)`) so
either commas, semicolons, or spaces work. All listed addresses receive the
same email in a single SMTP send (one `sendmail` call with the recipient list,
and the `To:` header shows all addresses).

To change recipients, edit `NOTIFY_TO` in `scripts/.env`. If `.env` is missing
or the password is empty, the run logs `email skipped` and still completes.
`--no-email` skips the email for a single manual run.

### Caveats
- Tasks run only while the user is logged in (Windows default). To run when
  logged out, open `taskschd.msc` and tick *Run whether user is logged on or
  not* (requires the Windows password).
- If a workbook is open in Excel during a scheduled run, the write step fails
  with a PermissionError — keep the `.xlsx` files closed overnight.

## Per-domain rules

- **AD** — rank by I-AUROC (tiebreak P-AUROC). 5 sub-areas: Anomaly Detection /
  Unsupervised / One-Class / Supervised / Graph.
- **OD** — each dataset has its **own** primary metric (COCO AP, mAP@0.5,
  Streaming AP, S-measure, novel AP, MAE…). The `_od_extract()` function in
  `update_papers.py` handles this. 5 sub-areas incl. 3D OD.
- **CLS** — rank by Top-1 (tiebreak Top-5). 5 sub-areas incl. Semi-Supervised
  and Small-Data. Never mix different SSL label budgets in one ranking.
- **AS** — domain disabled on the website ("待開發中"); do not surface.
- Different datasets and different evaluation protocols are **never** mixed in
  one ranking. Flag incomparable rows in the notes column.

## References

- `references/source_access.md` — how to search/export from each of the 6 sources
- `references/excel_schema.md`  — exact column layout for every workbook
