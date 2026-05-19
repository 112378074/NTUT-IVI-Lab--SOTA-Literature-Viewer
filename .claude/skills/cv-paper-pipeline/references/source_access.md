# Source Access Guide

How to pull AD / OD / Classification papers from each of the six sources.
Two categories: **programmatic** (the skill can fetch automatically) and
**login-gated** (manual metadata export only).

---

## 1. arXiv — programmatic ✅

- API: `http://export.arxiv.org/api/query`
- Already integrated in `scripts/update_papers.py`.
- Rate limit: ~1 request / 3 s. The script uses 5 s spacing + retry/backoff.
- Covers preprints; most CVPR/ICCV/ECCV papers also appear here first.

Run: `python scripts\update_papers.py`

---

## 2. CVF Open Access — programmatic ✅

- URL pattern: `https://openaccess.thecvf.com/{CONF}{YEAR}?day=all`
  (CONF ∈ CVPR / ICCV / ECCV / WACV)
- Fully public HTML, no login.
- Scraper: `.claude/skills/cv-paper-pipeline/scripts/fetch_cvf.py`

Run: `python .claude\skills\cv-paper-pipeline\scripts\fetch_cvf.py --conf CVPR --year 2025`

Output `cvf_results.csv` → review → merge into the xlsx workbooks.

---

## 3. IEEE Xplore — login-gated ⚠️

No free bulk API. The IEEE Xplore Search API requires a paid key. Manual path:

1. Go to <https://ieeexplore.ieee.org/> (on the school VPN for full access).
2. Search box — paste one query at a time:
   - `("anomaly detection" OR "defect detection") AND ("MVTec" OR "VisA")`
   - `("object detection") AND ("COCO" OR "LVIS") AND (2024 OR 2025 OR 2026)`
   - `("image classification") AND ("ImageNet" OR "CIFAR") AND ("semi-supervised" OR "few-shot")`
3. Filter: Year 2021-2027, Content Type = Conferences + Journals.
4. Select results → **Export** → **Citations** → format **CSV** → download.
5. Save as `sources/ieee_export.csv`, then hand the path to the skill.

---

## 4. Web of Science — login-gated ⚠️

Requires institutional subscription. Manual path:

1. <https://www.webofscience.com/> (school VPN / login).
2. Use the Advanced Search with `TS=` topic queries:
   - `TS=(("anomaly detection" OR "defect detection") AND ("MVTec" OR "VisA" OR "MPDD"))`
   - `TS=(("object detection") AND ("COCO" OR "PASCAL VOC" OR "LVIS")) AND PY=(2021-2027)`
   - `TS=(("image classification") AND ("few-shot" OR "fine-grained" OR "semi-supervised"))`
3. **Export** → **Tab-delimited / Excel** → "Records 1-N" → Full Record.
4. Save as `sources/wos_export.csv` (or `.txt`).

---

## 5. ScienceDirect — login-gated ⚠️

Elsevier; full text needs subscription. Manual path:

1. <https://www.sciencedirect.com/> (school VPN).
2. Search, e.g. `"industrial anomaly detection" surface defect`.
3. Filter Years 2021-2027, Article type = Research articles.
4. Select → **Export** → **Export citation to RIS / BibTeX / CSV**.
5. Save as `sources/sciencedirect_export.ris` (or `.csv`).

---

## 6. Scopus — login-gated ⚠️

Elsevier citation database; needs subscription or a paid Scopus API key.
Manual path:

1. <https://www.scopus.com/> (school VPN).
2. Document search with `TITLE-ABS-KEY(...)`:
   - `TITLE-ABS-KEY("anomaly detection" AND ("MVTec" OR "VisA")) AND PUBYEAR > 2020`
   - `TITLE-ABS-KEY("object detection" AND "COCO") AND PUBYEAR > 2020`
3. Select all → **Export** → **CSV** → include Citation information +
   Bibliographical information + Abstract.
4. Save as `sources/scopus_export.csv`.

---

## Importing a manual export

Once a CSV/RIS file is in `sources/`, the skill parses it with pandas, maps the
columns to the workbook schema (`references/excel_schema.md`), deduplicates
against existing rows by **title / DOI / arXiv ID**, and appends new rows.

Common column name mappings:

| Workbook field | IEEE CSV | WoS export | Scopus CSV |
| --- | --- | --- | --- |
| title    | `Document Title` | `Article Title` (`TI`) | `Title` |
| authors  | `Authors`        | `Authors` (`AU`)       | `Authors` |
| venue    | `Publication Title` | `Source Title` (`SO`) | `Source title` |
| year     | `Publication Year` | `Publication Year` (`PY`) | `Year` |
| doi      | `DOI`            | `DOI` (`DI`)           | `DOI` |
| link     | `PDF Link` / `Document URL` | derive from DOI | `Link` |

Metrics are **not** in these exports — they must be read from the paper and
verified before being entered (see the metrics policy in `SKILL.md`).
