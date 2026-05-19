# Excel Workbook Schema

Exact column layout for each workbook. **Column order is fixed** — the
`regenerate_*_json()` functions read by header name, so headers must match.

---

## AD — `AnomalyDetection_Papers_Summary_v10_20260425.xlsx`

Sheets: `說明 Notes`, `總覽 All Papers`, `MVTec AD`, `MVTec LOCO`,
`MVTec AD2`, `VisA`, `MPDD`, `BTAD`, `MVTec 3D`.

Per-dataset sheets — **header on row 3**, data from row 4. 15 columns:

| # | Column | Notes |
|---|---|---|
| 1 | 資料集 | Dataset name |
| 2 | 類別 | Method category (基於重構 / 基於 NF / 基於表徵 / 基於資料擴增 / 基於擴散) |
| 3 | 方法 | Method name |
| 4 | 作者 | Authors |
| 5 | 發表 | Venue |
| 6 | 年月 | YYYY-MM |
| 7 | 狀態 | 已發表 / 預印本 / 在審 |
| 8 | I-AUROC | Image-level AUROC (%) |
| 9 | P-AUROC | Pixel-level AUROC (%) |
| 10 | P-AP | Pixel AP (%) |
| 11 | P-PRO | Per-Region Overlap (%) |
| 12 | FPS | Inference speed |
| 13 | 備註(特色/based) | Notes / features |
| 14 | 連結 | Paper URL |
| 15 | GitHub | Repo URL or N/A |

`總覽 All Papers` uses the same 15 columns, **header on row 1**.
Ranking: I-AUROC desc, tiebreak P-AUROC desc.

---

## OD — `Object_Detection_Papers_Ranking_2021_2026.xlsx`

Sheets: `OD all papers`, `Index`, then one per dataset (35 dataset sheets).

`OD all papers` — header row 1, 8 columns:
`類別, 方法, 作者, 發表, 年月, 備註(特色/based), 連結, GitHub`

Per-dataset sheets — header row 1, 13 columns:

| # | Column | Notes |
|---|---|---|
| 1 | 類別 | General OD / Real-Time OD / RGB Salient OD / Few-Shot OD |
| 2 | 方法 | Method name |
| 3 | 作者 | Authors |
| 4 | 發表 | Venue |
| 5 | 年月 | YYYY-MM |
| 6 | 狀態 | Published / Preprint / … |
| 7 | mAP | Numeric COCO-style mAP (blank when N/A) |
| 8 | AP | Free-text metric string, e.g. `S 0.939 / Fβ 0.940 / MAE 0.014`, `sAP 37.8`, `AP 14.1` |
| 9 | FPS | Inference speed |
| 10 | params | Parameter count |
| 11 | 備註(特色/based) | Notes |
| 12 | 連結 | Paper URL |
| 13 | GitHub | Repo URL or N/A |

`Index` sheet maps each dataset → category + primary metric. The `資料集 Sheet`
column holds `=HYPERLINK("#'sheet'!A1","sheet")` formulas — read the display
label, not the formula.

**Per-dataset primary metric** (handled by `_od_extract()` in
`update_papers.py`): COCO family → mAP column; PASCAL VOC → AP50 in mAP;
Argoverse-HD → `sAP` parsed from the AP string; SOD datasets (DUTS-TE, ECSSD,
…) → `S-measure` from the AP string; shadow datasets (SBU-Refine, ISTD) →
`MAE` (lower is better); Few-Shot COCO → `novel AP` from the AP string.

---

## CLS — `Image_Classification_Papers_Ranking_2021_2026.xlsx`

Sheets: `Classification all papers`, then one per dataset (34 dataset sheets).

All sheets — header row 1, 15 columns:

| # | Column | Notes |
|---|---|---|
| 1 | 類別 | General / Fine-Grained / Long-Tailed / Few-Shot / Semi-Supervised / Small Data Image Classification-based |
| 2 | 方法 | Method name |
| 3 | 作者 | Authors |
| 4 | 發表 | Venue |
| 5 | 年月 | YYYY-MM |
| 6 | 狀態 | Published / Preprint |
| 7 | Top-1 | Top-1 accuracy (%) |
| 8 | Top-5 | Top-5 accuracy (%) |
| 9 | Acc | Generic accuracy when Top-1/5 not split |
| 10 | F1 | F1 score |
| 11 | Params | Parameter count |
| 12 | FLOPs | FLOPs |
| 13 | 備註(特色/based) | Notes — record SSL label budget here (e.g. "40 labels") |
| 14 | 連結 | Paper URL |
| 15 | GitHub | Repo URL or N/A |

Ranking: Top-1 desc, tiebreak Top-5 desc. Never mix different SSL label
budgets in one ranking.

`cls_supplemental.json` — hand-curated rows merged on top of the xlsx by
`regenerate_cls_json()`. Use it when the xlsx is locked (open in Excel). Each
row needs: `category, method, authors, venue, date, status, dataset, top1,
top5, acc, f1, params, flops, notes, link, github, source`.

---

## AS — `Anomaly_Synthesis_Papers_Benchmark_2021_2026.xlsx`

Domain disabled on the website. Sheets: `AS all papers`, then per-dataset
(18 sheets). 20 columns:
`類別, 方法, 作者, 發表, 年月, 狀態, 合成類型, 任務設定, 資料集,
Split/Protocol, I-AUROC, P-AUROC, AUPRO/PRO, AP/AUPRC, F1/Dice/IoU,
合成品質指標, Backbone/Detector, 備註(特色/可比性), 連結, GitHub`

---

## Adding rows safely

1. Match the header names exactly; keep the fixed column order.
2. Copy cell styling from the previous row (font / border / fill) so the new
   row blends in — see `style_from()` / `append_ad_row()` in `update_papers.py`.
3. Append to **both** the per-dataset sheet and the "all papers" sheet.
4. Leave a metric blank rather than guessing. Verified numbers only.
5. After editing, run the regenerate step (SKILL.md Step 3) and check the
   counts before committing.
