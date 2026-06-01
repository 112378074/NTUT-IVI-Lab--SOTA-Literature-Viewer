# Journal-finder spec schema

The agent produces a JSON file, then runs
`scripts/build_journal_xlsx.py --spec <file> --outdir <dir>`.

## Top-level fields

| key | type | required | meaning |
| --- | --- | --- | --- |
| `title` | string | recommended | Title shown on the 說明 sheet. |
| `output` | string | recommended | Output filename (relative to `--outdir`). Default `journal_list.xlsx`. |
| `date` | string | recommended | Compilation date, e.g. `2026-06-01`. |
| `relevance_label` | string | recommended | Header for the last column, e.g. `對 IVAD 關聯`. Default `關聯`. |
| `domains` | array | required | Domain buckets; each `{ "name": str, "color": "RRGGBB" }`. `color` optional (auto-cycled). |
| `journals` | array | required | The journal rows (see below). |
| `verified_if` | array of string | optional | One line per verified IF, with its source, listed on the 說明 sheet. |
| `sources` | array of [label, url] | optional | Source links rendered on the 說明 sheet. |
| `notes` | array of [key, value] | optional | Extra key/value note rows on the 說明 sheet. |

## `journals[]` item

| field | required | meaning |
| --- | --- | --- |
| `domain` | yes | Must match a `domains[].name` (controls row color + per-domain sheet). |
| `tier` | yes | `Tier1` / `Tier2` / `Tier2/3` / `Tier3` / `子領域選用` (free text). |
| `name` | yes | Full journal name. |
| `abbrev` | no | Common abbreviation (e.g. `IEEE TPAMI`). |
| `publisher` | yes | IEEE / Springer / Elsevier / MDPI / Taylor & Francis / IET / Emerald / … |
| `if` | **verified only** | JCR impact factor as a string (e.g. `"7.6"`). **Leave `""` if not verified this run.** |
| `quartile` | high-confidence only | `Q1`…`Q4`. Leave `""` if unsure. |
| `issn` | no | Common print ISSN for filing reference. Leave `""` if unsure. |
| `scope` | yes | One-line scope in the user's language. |
| `relevance` | yes | `◎◎` best-fit / `◎` high / `○` medium / `△` marginal / `選用` sub-area. |

## Anti-fabrication rule (must follow)

`if`, `quartile`, and `issn` are facts. Fill them **only** when verified from a
cited web source during this run, and record the source in `verified_if` /
`sources`. Otherwise set the field to `""`. Never guess or carry over a
remembered number. Blank cells are expected and correct.

## Output

The builder emits:
- `全部期刊` — combined table (rows colored by domain), frozen header row, AutoFilter.
- one sheet per domain (sheet title sanitized to ≤31 chars, illegal chars stripped).
- `說明` — notes + anti-fabrication rules + `verified_if` list + `sources` links.

Column order is fixed:
`領域 / Tier / 期刊全名 / 縮寫 / 出版社 / IF (JCR) / 分區 / ISSN / 範疇 / <relevance_label>`.
