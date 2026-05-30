# Classify-&-Verify workflow (MANDATORY before any leaderboard insert)

**Hard rule:** a newly-found OD or AD paper is **NEVER** auto-inserted into a
standard sortable leaderboard. It enters as catalog-only / blank-metric until it
passes the workflow below. Auto-fetched rows (arXiv/CVF scheduled run) land in
the "All Papers" catalog with **no metric value** and status **⏳ Pending visual
table read** — they only become leaderboard rows after a human/agent verifies
them with this procedure.

> If uncertain whether a number belongs to the method / dataset / split /
> protocol / metric, do **not** put it in the standard leaderboard. Mark it
> `⏳ Pending visual table read` or `⚠️ No authoritative source`.

---

## Status-label vocabulary (col 6 「狀態」)

These drive `_od_comparable()` / `_od_protocol()` in `scripts/update_papers.py`
and the colored badges + protocol grouping in `index.html`.

| Status label | Comparable group | Protocol | Badge |
| --- | --- | --- | --- |
| `✅ Standard verified` | Standard | Fully-supervised closed-set | green |
| `✅ Cross-paper verified` | Standard | Fully-supervised closed-set | green |
| `🟣 Open-vocabulary verified` | Non-standard | Open-vocabulary | purple |
| `🟠 Zero-shot verified` | Non-standard | Zero-shot | orange |
| `🟢 Few-shot verified` | Non-standard | Few-shot | green-cyan |
| `🟡 Non-standard protocol verified` | Non-standard | Non-standard protocol | yellow |
| `⚠️ Different dataset` / `Different metric` / `Misplaced task` | Non-standard | — | amber |
| `⚠️ No authoritative source` | Non-standard (no rank) | — | amber, **blank score** |
| `⏳ Pending visual table read` | Pending (bottom, no rank) | — | grey |

Rule: **never** label a verified non-standard-protocol result as "unverified".
If the source is checked, use the colored 🟣/🟠/🟢/🟡 label, not ⚠️.

---

## Filling a sheet to ≥10 papers (do NOT pad by mixing protocols)

Goal: every dataset page should hold **≥10 papers, all year 2020–2026, all
source-verified**; new papers keep slotting in and re-sorting.

**Hard constraint — never merge different protocols into the same Standard
Leaderboard just to reach 10 rows.** Reaching "10 on the page" may legitimately
*include* protocol groups, but the **Standard Leaderboard itself only counts
same-dataset + same-split + same-standard-protocol + same-primary-metric +
source-verified rows.**

- A same-dataset paper with a **different protocol** (open-vocab / zero-shot /
  few-shot / long-tail / domain-adapt / robustness / 3D / logical / seg-only …)
  stays on the page but **only under its correct Protocol Group**, with
  **Protocol, Primary Metric, Source Note, and Comparable Group** all labeled
  (colored 🟣/🟠/🟢/🟡 status). It does **not** enter the standard ranking.
- A paper on a **different dataset** goes to that dataset's page (or a new sheet),
  never padded onto this one.
- If genuine **same-protocol standard** papers number fewer than 10, that is
  acceptable — **flag it, never pad** with mismatched-protocol or unsourced rows.
  (e.g., a niche benchmark may only have 6 fully-supervised papers; report "6
  Standard + N protocol-group + needs-more" rather than inventing 4.)
- AD equivalent: the Standard AD board counts only standard-unsupervised + same
  metric (rank by I-AUROC); zero-shot/few-shot/continual/3D-multimodal/logical/
  seg-only rows stay on the dataset page under their own labeled protocol group.

---

## OD workflow (9 steps)

### Step 1 — Identify the actual task type
Read title + abstract + experiments + result tables. Classify into exactly one:
Standard 2D closed-set OD · Real-Time OD · 3D/BEV OD · RGB Salient/DIS/Co-SOD ·
Few-Shot OD · Cross-Domain Few-Shot OD · Open-Vocabulary OD · Zero-Shot OD ·
Open-World OD · Long-Tail OD · Domain Adaptation/Generalization OD ·
Robustness/Security/Attack OD · Referring/Grounded/MLLM Detection ·
Oriented/Remote-Sensing OD · Other specialized OD.
**Do not** call it Standard OD just because it says "object detection" — confirm
the real task + eval protocol.

### Step 2 — Identify the exact dataset AND split
Record exactly: COCO test-dev / val2017 / minival / 2017 val · PASCAL VOC 2007
test · LVIS v1.0 val · LVIS minival · CrowdHuman · Waymo 2D · CPPE-5 · OV-COCO ·
OV-LVIS · DOTA / DIOR-R / HRSC2016 · KITTI / nuScenes / Waymo 3D · SOD/Co-SOD.
**Never mix val / minival / test-dev.** A LVIS-minival number does NOT go on the
LVIS v1.0 val sheet unless the sheet explicitly allows minival and the note says so.

### Step 3 — Identify the protocol
Fully-supervised closed-set · Open-vocabulary · Zero-shot · Few-shot ·
Cross-domain few-shot · Open-world · Long-tail/rare-class · Domain adaptation ·
Domain generalization · Robustness/corruption/attack · Incremental · Oriented
bbox · 3D/BEV · SOD/DIS/Co-SOD · Referring/grounding.
Same dataset + different protocol → keep on the same dataset page **but label the
protocol** (colored status). Never merge into the standard ranking unlabeled.

### Step 4 — Identify the primary metric (do NOT convert metrics)
- COCO: **AP** primary; AP50/AP75/APS/APM/APL secondary
- PASCAL VOC: **mAP@0.5** · LVIS standard: **AP** · CrowdHuman: **AP** primary, MR⁻² secondary
- Waymo 2D: **AP/L1** primary, AP/L2 secondary · CPPE-5: **AP50** (dataset-paper main metric)
- COCO-O: **COCO-O mAP** (robustness) — **never** infer from normal COCO AP
- Open-Vocab: APN50 / APr / APm / AP_all (per benchmark) · Few-Shot: novel AP / nAP / k-shot AP
- SOD/DIS/Co-SOD: S-measure / F-measure / E-measure / MAE
- Open-World: U-Recall / WI / A-OSE / mAP · Domain-Adapt: target-domain AP
- Robustness/Security: robust AP / mAP-drop / ASR · Oriented: oriented mAP/AP50 · 3D/BEV: 3D AP / BEV AP / NDS
- **Never** treat S-measure / novel AP / U-Recall / ASR as standard COCO AP.

### Step 5 — Decide where it goes
- **A. Same dataset+split+standard protocol+primary metric** → Standard Leaderboard, `✅ Standard verified`, sort by primary metric.
- **B. Same dataset, different protocol (OV/zero-shot/few-shot/…)** → same dataset page, under the correct Protocol Group, status `🟣/🟠/🟢/🟡 … verified`. Not "unverified". Not in the standard ranking.
- **C. Different dataset** → move to that dataset's page; if it doesn't exist, create a new dataset sheet under the right OD sub-area. Do **not** force it onto COCO/VOC/LVIS.
- **D. Different metric** → keep the real metric, add to the task-specific leaderboard; mark Different metric / Non-standard if shown on a related page.
- **E. No authoritative source for the score** → status `⚠️ No authoritative source`, **score blank/"—"**, never in a sortable rank; keep only if useful as a related paper.
- **F. Placeholder / duplicate / invalid link / unidentifiable** → remove, or keep only in a needs-reference queue; never on the public leaderboard.

### Step 6 — Source verification (every numeric score)
Note must record: paper name / arXiv ID · table number or page · dataset split ·
metric name · original-paper-verified vs cross-paper-verified. Examples:
- `✅ Standard verified — COCO val AP, Table 2 of the paper.`
- `✅ Cross-paper verified — reported by UniNet CVPR25 Table 1(b), not the original method paper.`
- `🟣 Open-vocabulary verified — OV-LVIS AP_rare, Table 3; not comparable to fully-supervised LVIS AP.`
- `⚠️ No authoritative source — value not found in original paper or benchmark table.`

### Step 7 — Sorting (handled by index.html)
Default: (1) Standard comparable rows first, by dataset primary metric → (2)
protocol-specific groups, each by its own metric → (3) no-source/pending rows at
the bottom, **no numeric rank**. Same-dataset-different-protocol rows are grouped
& labeled, never deleted. The `🔀 全部 protocol 一起比較` toggle merges supported rows
into one primary-metric overview (excludes ⚠️ No-source / ⏳ Pending).

### Step 8 — Required fields per entry
Method · Paper title · Sub-area · Task Type · Dataset · Dataset split · Protocol ·
Primary Metric · Primary Score · Secondary Metrics · Comparable Group
(Standard / Protocol-specific / Related / No-source) · Verification Status ·
Source Type (original paper / official leaderboard / benchmark table / no source) ·
Source Note · Reason (if non-standard or excluded) · Paper link · GitHub · Year/Venue.
(In the workbook: col1 類別=task, col2 方法, col6 狀態=status label, col7 mAP=primary,
col8 AP="KEY=VALUE | …" breakdown, col11 備註=source note, col12 連結.)

### Step 9 — After a batch
1. Regenerate JSON/HTML (Step 3 of SKILL.md). 2. Re-sort affected pages.
3. Confirm standard vs protocol groups render correctly. 4. No value-bearing row
without a source note. 5. No verified row mislabeled "unverified". 6. No invalid
arXiv placeholder. 7. Preview the site (Preview MCP, check console = 0 errors).
8. Commit + push. 9. **Report**: new added · moved to other task/dataset · kept
as non-standard protocol · excluded no-source · final affected dataset counts ·
commit hash.

---

## AD workflow

Same gate: a new AD paper is **not** auto-inserted into the standard AD
leaderboard. First identify:
1. **Dataset** — MVTec AD · VisA · BTAD · MPDD · MVTec LOCO AD · MVTec 3D-AD · Real-IAD · MANTA …
2. **Protocol** — standard unsupervised · zero-shot · few-shot · continual · 3D/multimodal · logical AD · synthetic-anomaly training · segmentation-only.
3. **Metric** — Image AUROC · Pixel AUROC · Pixel AP · AU-PRO / PRO · accuracy · S-measure …
4. **Source** — original paper table · official benchmark · authoritative comparison table.

**Put into the Standard AD Leaderboard only if:** same dataset · same standard
(unsupervised) protocol · same metric · score is source-verified.

**Same dataset, different protocol** → keep on the same dataset page but label
clearly: `Zero-shot AD` · `Few-shot AD` · `Continual AD` · `3D / multimodal AD` ·
`Logical AD` · `Segmentation-only`. Use a verified label, **not** "unverified",
when the source is checked.

**Constraints:**
- Do not mix Image AUROC / Pixel AUROC / Pixel AP / PRO into one ranking unless the
  UI states which metric sorts. AD ranks by **I-AUROC** (tiebreak P-AUROC); other
  metrics are secondary columns, not a merged rank.
- A value with no authoritative source → not placed in any sortable leaderboard
  (blank score, `⚠️ No authoritative source`).
- Never fabricate numbers; only fill a metric verified from the paper's own table.

---

## Precedents (this repo) — examples of the rule in action
- **Focal-L DINO 58.4** (COCO test-dev) → deleted: FocalNet paper reports only FocalNet-H+DINO 64.3/64.4; 58.4 had no source.
- **COCO-O** → rebuilt from the benchmark's own Table 7 (49 tested detectors); RT-DETR/Co-DETR/etc. removed (the benchmark never tested them; no cross-paper COCO-O source).
- **PASCAL VOC 2007** → kept only InternImage-H 94.0 / DETReg 83.3 / Faster R-CNN 73.2 (real VOC sources); deleted DINO/YOLOv7/DETR/… community re-evals.
- **LVIS v1.0 val** → 7 fully-supervised Standard; 11 OV/zero-shot/few-shot kept on the same page as 🟣/🟠/🟢 protocol-verified; LVIS test-dev sheet removed (all rows were val/minival mislabels).
- **Argoverse-HD** → 4 byte-identical split sheets consolidated to 1; "To verify" reproductions deleted.
