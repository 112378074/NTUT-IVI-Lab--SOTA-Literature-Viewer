---
name: journal-finder
summary: "Find relevant academic journals for a research domain, verify their JCR impact factor / quartile, and export a styled Excel workbook."
description: >
  Given one or more research domains or topics (e.g. "影像 / 電腦視覺",
  "智慧製造", "industrial anomaly detection", "medical image analysis"),
  search the web for the most relevant academic journals, verify each
  journal's current JCR impact factor and quartile from cited sources,
  collect publisher / abbreviation / ISSN / scope, rate each journal's
  relevance to the user's research focus, and export a styled multi-sheet
  Excel workbook (combined sheet + per-domain sheets + a 說明/sources sheet).
  Strictly anti-fabrication: only impact-factor, quartile, and ISSN values
  verified from a cited source this run are filled; everything else is left
  blank with a note. Use this skill when the user asks to "find journals",
  "搜尋期刊", "整理期刊清單", "幫我找相關的期刊", or to build/refresh a
  journal target list for submission or literature search.
---

# Journal Finder（期刊搜尋）

End-to-end workflow: **parse domains → web-search journals → verify IF/quartile → build journal table → export styled .xlsx**.

```text
Research domain(s)
      │
      ▼
WebSearch  (top journals + JCR impact factor / quartile)
      │
      ▼
Verify gate  (only fill values confirmed from a cited source)
      │
      ▼
JSON spec  (journals[] + domains[] + sources[])
      │
      ▼
build_journal_xlsx.py  →  styled .xlsx (combined + per-domain + 說明)
```

## Core principle — never fabricate metrics

This mirrors the project's literature rules. A journal **name, publisher, abbreviation, and scope** may come from domain knowledge. But **impact factor, quartile, and ISSN are numeric/identifier facts** — fill them **only** when verified from a cited web source during this run.

- Verified value → fill the cell **and** add the source URL to `sources`.
- Not verified → leave the cell **blank** (`""`). Do **not** guess, interpolate, or carry over a remembered value.
- List every verified IF in `verified_if` so the 說明 sheet documents provenance.

It is correct and expected for many `if` / `quartile` / `issn` cells to be blank.

## Step 1 — Parse the requested domain(s)

Identify each distinct research area the user named. Each becomes a `domain` with a row-fill color. Examples:

- `影像 / 電腦視覺` (image / computer vision)
- `智慧製造` (smart / intelligent manufacturing, Industry 4.0)
- `industrial anomaly detection`, `medical image analysis`, `remote sensing`, …

If the user names a broad field, also decide a per-journal **relevance label** (the last column header), e.g. `對 IVAD 關聯`, `對研究關聯`. Relevance grades: `◎◎` best-fit submission/search target · `◎` high · `○` medium · `△` marginal / high-volume · `選用` sub-area optional.

## Step 2 — Search the web for journals + metrics

For each domain run a small set of `WebSearch` queries. Two query shapes work well:

1. **Discovery** — find the top journals and their metrics:
   ```
   top <domain> journals 2025 2026 impact factor list quartile
   ```
2. **Verification** — confirm a specific journal's IF/quartile, quoting the journal name:
   ```
   "<Journal Name>" impact factor 2025 quartile JCR
   ```

Good metric sources seen in practice: `journalmetrics.org`, `research.com`, `resurchify.com`, `bioxbio.com`, `scimagojr.com`, `wos-journal.info`, publisher pages. Cross-check when a value looks off; prefer the most recent JCR year. Record the URL you actually used.

Do not bypass logins/paywalls or bulk-download. Public metric-aggregator pages and publisher pages only.

## Step 3 — Build the journal table

Collect per journal:

| field | meaning | source |
| --- | --- | --- |
| `domain` | which domain bucket | from Step 1 |
| `tier` | rough standing: `Tier1` / `Tier2` / `Tier2/3` / `Tier3` / `子領域選用` | judgment from metrics + reputation |
| `name` | full journal name | knowledge |
| `abbrev` | common abbreviation (e.g. IEEE TPAMI) | knowledge |
| `publisher` | IEEE / Springer / Elsevier / MDPI / Taylor & Francis / … | knowledge |
| `if` | JCR impact factor — **verified only**, else `""` | web (Step 2) |
| `quartile` | `Q1`…`Q4` — high-confidence only, else `""` | web (Step 2) |
| `issn` | common print ISSN, for filing reference, else `""` | knowledge / publisher page |
| `scope` | one-line scope in the user's language | knowledge |
| `relevance` | `◎◎/◎/○/△/選用` vs the research focus | judgment |

Order rows by domain, then tier, then (verified) IF descending.

## Step 4 — Write the JSON spec and export

Write a JSON spec (see `references/spec_schema.md` and `references/sample_spec.json`) then run the builder:

```bash
python .claude\skills\journal-finder\scripts\build_journal_xlsx.py --spec <spec.json> --outdir "C:\Users\user\Desktop\Mypaper"
```

The builder produces:
- **全部期刊** — combined table (domain-colored rows), frozen header, AutoFilter.
- one sheet **per domain**.
- **說明** — notes, the anti-fabrication rules, the verified-IF list, and clickable source URLs.

Name the output descriptively, e.g. `期刊清單_影像CV_智慧製造.xlsx`, and save it to the project root unless the user says otherwise.

### Spec essentials

```json
{
  "title": "期刊清單_<domains>",
  "output": "期刊清單_<domains>.xlsx",
  "date": "YYYY-MM-DD",
  "relevance_label": "對 <focus> 關聯",
  "domains": [{"name": "影像 / 電腦視覺", "color": "E8F1FB"},
              {"name": "智慧製造", "color": "FCEFE3"}],
  "journals": [{"domain":"...","tier":"Tier1","name":"...","abbrev":"...",
                "publisher":"...","if":"7.6","quartile":"Q1","issn":"...",
                "scope":"...","relevance":"◎"}],
  "verified_if": ["Pattern Recognition 7.6 — https://..."],
  "sources": [["Pattern Recognition (IF 7.6, Q1)", "https://..."]]
}
```

Colors cycle automatically if omitted. Leave `if`/`quartile`/`issn` as `""` when unverified.

## Step 5 — Report

After building, report to the user:

1. Output file path + total journal count + per-domain counts.
2. Sheet structure (combined / per-domain / 說明).
3. Which IF/quartile values were **verified** and their sources.
4. Which cells were left blank (and that this is intentional, not an omission).
5. Offer follow-ups: fill remaining IFs by querying JCR/Scimago one-by-one, add an "official submission URL" column, merge with an existing list, or hand the list to the literature pipeline.

## Updating / extending an existing list

To add a domain or refresh metrics, reuse the prior spec JSON: append journals / refresh verified `if` values (re-search), then re-run the builder. Keep the anti-fabrication rule on every refresh — re-verify before changing a number.

## References

- `references/spec_schema.md` — full JSON field reference.
- `references/sample_spec.json` — a minimal working spec (image/CV + smart-manufacturing).
- `scripts/build_journal_xlsx.py` — the styled-workbook generator.
