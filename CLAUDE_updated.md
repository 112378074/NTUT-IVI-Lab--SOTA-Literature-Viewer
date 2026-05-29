# CLAUDE.md — Object Detection Paper Watcher & Dataset-wise SOTA Ranking Builder

**Version:** v3.0 — missed-paper prevention update  
**Purpose:** Build and maintain a 2021–2026 Object Detection literature review and dataset-wise ranking workbook.  
**Scope:** General Object Detection, Real-Time Object Detection, RGB Salient Object Detection, and Few-Shot Object Detection.  
**Hard exclusion:** 3D Object Detection and all LiDAR / BEV / camera-only 3D detection methods.

---

## 0. Why the previous search missed papers — root-cause notes

The previous search was incomplete for four main reasons. Future agents must explicitly avoid these failure modes.

1. **The search relied too heavily on existing leaderboard snapshots.**  
   Papers with Code and similar leaderboards are useful but often lag behind arXiv, official model documentation, GitHub releases, technical reports, and newly indexed 2026 preprints.

2. **The 2026 recency pass was not forced.**  
   The search used broad 2021–2026 queries, but it did not run a separate high-priority pass for `2026`, `last 30 days`, `last 90 days`, `new release`, `latest`, `arXiv 2026`, `GitHub 2026`, and official model documentation.

3. **Model-name aliases were not exhaustive.**  
   Recent model families can appear under several names, for example `YOLO26`, `YOLOv26`, `YOLO 26`, `Ultralytics YOLO26`, or repo-specific names. The agent must search all common aliases, not only one spelling.

4. **The output requirement did not enforce a per-dataset minimum.**  
   Every dataset worksheet must contain at least **10 ranked SOTA / candidate-SOTA models** whenever public results exist. If fewer than 10 strictly comparable SOTA rows are available, include the best available candidate rows and mark them as `Not directly comparable` or `To verify`.

---

## 1. Main task

You are an expert research assistant specializing in **Object Detection literature review and benchmark analysis**.

Search and organize Object Detection papers, preprints, official model releases, official documentation, and GitHub implementations from **2021-01-01 to 2026-12-31**, excluding all 3D Object Detection methods.

Focus only on the following four categories:

1. **General Object Detection-based**
2. **Real-Time Object Detection-based**
3. **RGB Salient Object Detection-based**
4. **Few-Shot Object Detection-based**

For every valid paper / method / model variant, extract the fixed columns:

```text
類別, 方法, 作者, 發表, 年月, 狀態, mAP, AP, FPS, params, 備註(特色/based), 連結, GitHub
```

---

## 2. Absolute output requirements

Produce the following three files:

### 2.1 Excel workbook

```text
Object_Detection_Papers_Ranking_2021_2026.xlsx
```

Workbook requirements:

- Include `OD all papers` overview sheet.
- Include one worksheet per dataset.
- Each dataset worksheet must contain at least **10 SOTA / candidate-SOTA model rows** whenever public results exist.
- Sort each dataset sheet by mAP or the dataset primary metric.
- Do **not** sort by FPS.
- Apply category row colors defined in Section 12.
- Freeze header row.
- Add filters.
- Wrap text.
- Use consistent column widths.
- Keep fixed column order.
- Do not mix different splits unless clearly marked.
- Flag all results that are not directly comparable.

### 2.2 Markdown report

```text
Object_Detection_Papers_Ranking_Report.md
```

Report must include:

1. Search date and search time zone.
2. Sources searched.
3. Total paper / model-variant count.
4. Counts by category.
5. Counts by year.
6. Top-10 methods per dataset.
7. 2021–2026 trends.
8. Category-specific trends.
9. 2026 latest-model watchlist.
10. Ten papers worth reading first.
11. Comparability notes.
12. Rows that need manual verification.

### 2.3 CSV backup

```text
Object_Detection_Papers_Ranking_2021_2026.csv
```

This CSV must be a backup of the `OD all papers` sheet with the same fixed columns.

---

## 3. Categories and classification rules

### 3.1 General Object Detection-based

Include general 2D object detection tasks:

- one-stage detector
- two-stage detector
- anchor-based detector
- anchor-free detector
- CNN-based detector
- Transformer / DETR-based detector
- hybrid object detector
- open-set / robust object detection if the main task remains 2D general OD

Representative methods include, but are not limited to:

- Faster R-CNN
- Mask R-CNN
- Cascade R-CNN
- RetinaNet
- FCOS
- CenterNet
- Sparse R-CNN
- DETR
- Deformable DETR
- DINO
- Co-DETR
- Relation-DETR
- EVA
- MaxViT-based detector
- InternImage-based detector
- ConvNeXt-based detector
- DINO-X / Grounding-DINO-style detectors, only when evaluated as 2D object detection

### 3.2 Real-Time Object Detection-based

Include detectors emphasizing real-time inference, FPS, latency, edge deployment, or low computational cost:

- real-time object detection
- lightweight object detection
- efficient object detection
- YOLO-based detector
- RT-DETR-based detector
- mobile / edge detector
- low-FLOPs detector
- NMS-free real-time detector

Representative methods include, but are not limited to:

- YOLO family: YOLOv5 / YOLOv6 / YOLOv7 / YOLOv8 / YOLOv9 / YOLOv10 / YOLOv11 / YOLOv12 / YOLOv13 / YOLO26 / YOLOv26
- YOLOX
- YOLO-Master
- DAMO-YOLO
- PP-YOLO / PP-YOLOE / PP-YOLOE+
- RT-DETR / RT-DETRv2 / RT-DETRv3 / RT-DETRv4
- D-FINE
- DEIM
- RF-DETR / RF-DETR 2, if evaluated as real-time 2D detection
- EfficientDet
- NanoDet
- MobileDet
- RTMDet

**Important alias rule:** Always search both `YOLO26` and `YOLOv26`, plus `YOLO 26`, `Ultralytics YOLO26`, and `Ultralytics YOLOv26`.

### 3.3 RGB Salient Object Detection-based

Include RGB image salient object detection / salient region detection:

- RGB salient object detection
- salient object detection
- saliency detection
- salient region detection
- high-resolution salient object detection
- shadow-aware salient detection if evaluated on RGB SOD datasets

Representative methods include, but are not limited to:

- BiRefNet
- InSPyReNet
- M3Net
- CPD
- CFDN
- U^2-Net
- PoolNet
- BASNet
- ICON
- VST
- EDN
- MINet

### 3.4 Few-Shot Object Detection-based

Include few-shot / low-shot object detection:

- few-shot object detection
- one-shot object detection
- k-shot object detection
- base-to-novel detection
- open-world / open-vocabulary few-shot detection
- meta-learning object detection
- transfer-learning few-shot detection

Representative methods include, but are not limited to:

- TFA
- DeFRCN
- FSCE
- Meta R-CNN
- FSRW
- MPSR
- CD-ViTO
- hANMCL
- UniFS
- DETReg
- Grounding DINO few-shot / ODinW setting

---

## 4. Hard exclusions

Exclude the following even if they use the phrase “object detection”:

- 3D Object Detection
- LiDAR-based 3D Detection
- BEV 3D Detection
- camera-only 3D Detection
- multi-modal 3D Detection
- point-cloud detection
- autonomous-driving 3D detection benchmarks such as KITTI 3D, nuScenes 3D, Waymo 3D
- medical detection, unless the dataset is explicitly listed in this specification
- segmentation-only methods, unless the task is RGB Salient Object Detection
- tracking-only papers without a detection benchmark

DOTA may be included only if the method is **2D oriented object detection**, not 3D detection.

---

## 5. Target datasets and dataset-specific ranking metrics

Create one worksheet per dataset. Each worksheet must contain **10 rows minimum** when public results exist. Dataset ranking must follow the dataset's own accepted metric and protocol. Do not force all dataset types into COCO-style mAP.

### 5.1 General Object Detection datasets

Required datasets include, but are not limited to:

1. COCO test-dev
2. COCO minival
3. COCO-O
4. PASCAL VOC 2007
5. COCO 2017 val
6. COCO 2017
7. CrowdHuman full body
8. GraZPEDWRI-DX
9. CPPE-5
10. Waymo 2D detection

Optional but recommended:

- MS COCO
- PASCAL VOC 2012
- LVIS
- Objects365
- Open Images
- WiderFace
- Cityscapes
- BDD100K

Primary ranking metrics:

```text
COCO-style box AP / mAP, high-to-low
```

Secondary / constrained metrics:

```text
AP50 / mAP@0.5 only when comparing the same dataset, same split, and same evaluation setting
```

Rules:

- COCO AP normally means AP averaged over IoU thresholds 0.50:0.95.
- AP50 must not be treated as equivalent to COCO AP.
- Different splits such as `test-dev`, `val`, `minival`, and custom splits must be separated or clearly marked.

### 5.2 Real-Time Object Detection datasets

Required datasets include, but are not limited to:

1. MS COCO
2. COCO RT
3. PASCAL VOC 2007 RT
4. Argoverse-HD
5. Argoverse-HD Full-Stack Val
6. Argoverse-HD Detection-Only Val
7. Argoverse-HD Detection-Only Test
8. Argoverse-HD Full-Stack Test

Optional but recommended:

- VisDrone
- UAVDT
- BDD100K
- AI-TOD
- TinyPerson
- DOTA, only when the method is 2D oriented detection

Primary ranking metrics:

```text
box AP / COCO AP / mAP, high-to-low
```

Efficiency metrics to record but not use as the default ranking key:

```text
FPS, latency, parameters, FLOPs, input size, hardware
```

Rules:

- FPS and latency are efficiency supplements, not the main rank criterion, unless the user explicitly asks for speed-based ranking.
- Always record the test hardware, input size, precision mode, batch size, and whether TensorRT / ONNX / TensorRT-FP16 / TensorRT-INT8 was used when available.
- Do not directly compare FPS across different hardware or deployment backends without marking the result as `Not directly comparable`.

### 5.3 RGB Salient Object Detection datasets

Required datasets include, but are not limited to:

1. DUTS-TE
2. DUT-OMRON
3. HKU-IS
4. ECSSD
5. PASCAL-S
6. SBU
7. HRSOD
8. UHRSD
9. DAVIS-S
10. ISTD
11. CAMO-FS

Optional but recommended:

- SOC
- THUR15K
- MSRA-B
- DUTS-TR
- SBU-Refine
- RGBD-SOD dataset only if RGB-only results are separately reported

Primary ranking metrics:

```text
S-measure, high-to-low
E-measure, high-to-low
F-measure, high-to-low
MAE, low-to-high
```

Rules:

- RGB Salient Object Detection must not be ranked by general object detection mAP unless the paper explicitly reports a detection-style mAP for the same SOD task.
- S-measure / E-measure / F-measure are better when higher.
- MAE is better when lower.
- If multiple SOD metrics are available, preserve all of them in the remarks or dataset-specific metric columns and rank by the benchmark's primary metric.

### 5.4 Few-Shot Object Detection datasets

Required datasets include, but are not limited to:

1. PASCAL VOC 2007 15+5 split
2. MS-COCO 1-shot
3. MS-COCO 5-shot
4. MS-COCO 10-shot
5. MS-COCO 30-shot
6. COCO 2017 FSOD
7. LVIS v1.0 val
8. LVIS v1.0 test-dev
9. ODinW-13
10. ODinW-35

Optional but recommended:

- PASCAL VOC few-shot split
- COCO few-shot novel split
- FSOD
- DIOR few-shot
- DOTA few-shot, only if the method is 2D detection

Primary ranking metrics:

```text
AP, AP50, novel AP, base AP
```

Rules:

- Shot setting is mandatory: `1-shot`, `5-shot`, `10-shot`, `30-shot`, or the paper's exact k-shot setting.
- 1-shot, 5-shot, 10-shot, and 30-shot results must not be mixed into one ranking.
- Novel AP and base AP must not be mixed as the same ranking criterion.
- Different split definitions must be separated or explicitly marked.
- If a dataset reports both base and novel results, rank primarily by the benchmark-specified novel-class metric unless the user requests otherwise.

## 6. Search workflow — mandatory multi-pass process

Do not perform only one broad search. Use all passes below.

### Pass A — Dataset leaderboard pass

For each dataset, search:

```text
"<dataset name>" "object detection" "AP" "mAP" "leaderboard"
"<dataset name>" "Papers with Code" "object detection"
"<dataset name>" "state of the art" "object detection" "AP"
"<dataset name>" "benchmark" "object detection" "2026"
```

Collect the Top-10 available methods for each dataset.

### Pass B — 2026 latest-model pass

Run this pass before finalizing `OD all papers`.

```text
"object detection" "2026" "COCO" "AP" "arXiv"
"real-time object detection" "2026" "COCO" "FPS" "arXiv"
"latest object detector" "2026" "COCO" "AP"
"SOTA object detection" "2026" "COCO" "AP"
"YOLO26" "COCO" "mAP" "FPS"
"YOLOv26" "COCO" "mAP" "FPS"
"YOLO 26" "COCO" "mAP" "FPS"
"Ultralytics YOLO26" "COCO" "mAP" "FPS"
"YOLO-Master" "real-time object detection" "COCO" "AP"
"D-FINE" "COCO" "AP" "FPS"
"DEIM" "COCO" "AP" "FPS"
"RF-DETR" "COCO" "AP" "FPS"
"RF-DETR 2" "COCO" "AP" "FPS"
```

Any valid 2025–2026 method found here must be added to `OD all papers` even if full dataset-wise metrics are not yet available. Mark incomplete rows as `To verify`.

### Pass C — Model-family alias pass

Search every important model family using aliases:

```text
YOLO26 OR YOLOv26 OR "YOLO 26" OR "Ultralytics YOLO26"
YOLOv13 OR YOLO13 OR YOLO 13
YOLOv12 OR YOLO12 OR YOLO 12
RT-DETRv4 OR RTDETRv4 OR RT-DETR v4
D-FINE OR DFINE
DEIM object detection
RF-DETR OR RFDETR
Co-DETR OR CoDETR
Relation-DETR OR Relation DETR
BiRefNet salient object detection
CD-ViTO few-shot object detection
```

### Pass D — Source-specific search pass

Search these source types separately:

#### arXiv / cs.CV

```text
site:arxiv.org/abs "object detection" "COCO" "AP" "2026"
site:arxiv.org/abs "real-time object detection" "COCO" "FPS" "2026"
site:arxiv.org/abs "few-shot object detection" "COCO" "novel AP" "2026"
site:arxiv.org/abs "salient object detection" "DUTS-TE" "2026"
site:arxiv.org/abs "YOLO26"
site:arxiv.org/abs "YOLOv26"
```

#### CVF Open Access

```text
site:openaccess.thecvf.com "object detection" "COCO" "AP" "2026"
site:openaccess.thecvf.com "real-time object detection" "FPS" "COCO"
site:openaccess.thecvf.com "few-shot object detection" "COCO"
site:openaccess.thecvf.com "salient object detection" "DUTS-TE"
```

#### GitHub

```text
site:github.com "YOLO26" "COCO" "mAP"
site:github.com "YOLOv26" "object detection"
site:github.com "real-time object detection" "COCO" "FPS" "2026"
site:github.com "D-FINE" "COCO" "AP"
site:github.com "DEIM" "COCO" "AP"
site:github.com "RF-DETR" "COCO" "AP"
site:github.com "few-shot object detection" "COCO" "novel AP"
site:github.com "salient object detection" "DUTS-TE"
```

#### Official documentation

```text
site:docs.ultralytics.com "YOLO26"
site:docs.ultralytics.com "YOLOv26"
site:docs.ultralytics.com/models/yolo26
site:roboflow.com "YOLO26" "object detection"
site:learnopencv.com "YOLOv26" "object detection"
```

#### IEEE Xplore

```text
("object detection" OR "real-time object detection" OR "few-shot object detection" OR "salient object detection")
AND ("COCO" OR "PASCAL VOC" OR "DUTS-TE" OR "LVIS" OR "ODinW")
AND ("mAP" OR "AP" OR "FPS")
AND (2021 OR 2022 OR 2023 OR 2024 OR 2025 OR 2026)
```

#### Web of Science

```text
TS=(("object detection" OR "real-time object detection" OR "few-shot object detection" OR "salient object detection")
AND ("COCO" OR "PASCAL VOC" OR "LVIS" OR "DUTS-TE" OR "ODinW")
AND ("mAP" OR "AP" OR "FPS"))
AND PY=(2021 OR 2022 OR 2023 OR 2024 OR 2025 OR 2026)
```

### Pass E — Completion pass

Before writing the final Excel:

- Check that every required dataset worksheet has at least 10 rows.
- Check that every 2026 method found in Pass B appears in `OD all papers`.
- Check that YOLO26 / YOLOv26 search results were explicitly reviewed.
- Check that each row has a source link whenever possible.
- Check that GitHub is included when available.
- Check that rows without verified metrics are marked `To verify`.

### Pass F — Comparability pass

For every dataset worksheet:

- Do not mix test-dev, val, minival, and custom split without explicit marking.
- Do not compare single-scale and multi-scale results as identical unless the paper reports them under the same protocol.
- Do not treat AP50 as COCO mAP.
- Do not rank by FPS.
- Do not mix official leaderboard results with third-party reproductions without marking source type.

---

## 7. Ranking and sorting rules

### 7.1 `OD all papers` overview sheet

The `OD all papers` sheet is a chronological literature overview, not a benchmark leaderboard.

Sort `OD all papers` by:

```text
年月, newest-to-oldest
```

Required overview fields:

```text
類別, 方法, 作者, 發表, 年月, 備註, 連結, GitHub
```

Rules:

- Include every valid paper / model / official release found during the search, even if benchmark metrics are incomplete.
- Prefer `YYYY-MM` for `年月`. If only the year is known, use `YYYY-00` and mark the date as approximate in `備註`.
- `連結` and `GitHub` must be clickable hyperlinks in Excel and Markdown outputs.
- If GitHub is unavailable, write `N/A`; do not fabricate a repository.

### 7.2 Dataset worksheets

Each dataset worksheet is a benchmark ranking table. Sort each dataset sheet by that dataset's own primary metric.

General rule:

```text
Dataset primary metric, best-to-worst
```

Examples:

- COCO-style Object Detection: box AP / COCO AP / mAP, high-to-low.
- AP50-only datasets: AP50 / mAP@0.5, high-to-low, only within the same dataset and protocol.
- RGB Salient Object Detection: S-measure / E-measure / F-measure high-to-low, or MAE low-to-high.
- Few-Shot Object Detection: novel AP, AP, AP50, or benchmark-specified primary metric, with shot setting and split kept separate.

Rules:

- Do not sort dataset worksheets by FPS unless the user explicitly requests speed ranking.
- Do not mix validation, minival, test-dev, test, and custom splits without explicit marking.
- Do not mix single-scale, multi-scale, TTA, different training data, or different input sizes as if they were identical.
- Every dataset sheet should make the ranking metric visible, either through the `mAP` / `AP` columns or through a clear dataset-specific note in `備註(特色/based)`.

### 7.3 General OD and Real-Time OD

Sort each dataset worksheet by:

```text
mAP / AP from high to low
```

Use the dataset's primary metric. For COCO-style benchmarks, AP normally means COCO AP averaged over IoU thresholds 0.50:0.95.

Do not use AP50 as COCO mAP unless the paper explicitly defines AP50 as the main metric for that dataset.

### 7.4 RGB Salient Object Detection

RGB SOD often does not use mAP. If mAP is unavailable:

- Put `N/A` in the `mAP` column.
- Put the primary SOD metric in the `AP` column or in a dataset-specific metric note.
- Sort by the dataset's primary SOD metric.

Metric priority when the dataset does not specify one:

```text
S-measure high-to-low > E-measure high-to-low > F-measure high-to-low > MAE low-to-high
```

Add this note to remarks:

```text
RGB SOD uses S-measure / E-measure / F-measure / MAE instead of COCO AP; not directly comparable with general object detection mAP.
```

### 7.5 Few-Shot Object Detection

If base AP, novel AP, and all AP are all reported, prioritize:

```text
novel AP or benchmark-specified primary AP
```

Always record the shot setting:

```text
1-shot / 5-shot / 10-shot / 30-shot / exact k-shot setting
```

Never rank different shot settings or different split definitions in the same ordered list.

## 8. Fixed output columns

Every worksheet must use the same fixed columns:

| 欄位 | 說明 |
|---|---|
| 類別 | General OD / Real-Time OD / RGB Salient OD / Few-Shot OD |
| 方法 | Model or method name; separate model variants into different rows |
| 作者 | Paper authors |
| 發表 | CVPR / ICCV / ECCV / WACV / IEEE / arXiv / OpenReview / GitHub / official docs |
| 年月 | YYYY-MM |
| 狀態 | Published / Accepted / arXiv / Preprint / GitHub / Official Docs / Reproduced / To verify |
| mAP | Main mAP or AP number; use N/A if unavailable |
| AP | AP / AP50 / AP75 / novel AP / dataset-specific metric |
| FPS | Inference speed; use N/A if unavailable |
| params | Parameter count; use N/A if unavailable |
| 備註(特色/based) | Features, backbone, neck, head, based type, source credibility, SOTA status, comparability notes |
| 連結 | Paper / arXiv / CVF / IEEE / OpenReview / official docs / Papers with Code link |
| GitHub | Official or unofficial GitHub link |

Fixed order:

```text
類別, 方法, 作者, 發表, 年月, 狀態, mAP, AP, FPS, params, 備註(特色/based), 連結, GitHub
```

Hyperlink rules:

- In Excel, `連結` and `GitHub` must be inserted as clickable hyperlinks, not plain text only.
- In Markdown, `連結` and `GitHub` must use Markdown hyperlink syntax.
- Use the paper / arXiv / CVF / IEEE / OpenReview / official documentation URL for `連結`.
- Use the official GitHub repository whenever available; otherwise use the most credible implementation and mark it as unofficial in `備註(特色/based)`.
- If no reliable GitHub repository exists, write `N/A`.

---

## 9. Source credibility labels

Every row must include one of these labels in `備註(特色/based)`:

- Official leaderboard
- Papers with Code leaderboard
- Paper-reported result
- Official GitHub result
- Official documentation
- Third-party reproduction
- Unverified arXiv claim
- To verify

If results are not directly comparable, add:

```text
Not directly comparable: different training data / input size / test-time augmentation / split / metric definition.
```

---

## 10. OD all papers sheet rules

The `OD all papers` worksheet is not a dataset ranking table. It is the master chronological paper / method index and must be sorted by `年月` from newest to oldest.

For this sheet, the required concise view is:

```text
類別, 方法, 作者, 發表, 年月, 備註, 連結, GitHub
```

The `OD all papers` worksheet is not just a duplicate of dataset sheets. It must also include:

- Newly released 2025–2026 methods even if dataset-wise rows are incomplete.
- Methods found in official documentation or GitHub release pages.
- Methods without GitHub if they are important.
- Preprints without official publication if they are relevant.
- Multiple model variants as separate rows.
- Rows with incomplete metrics, clearly marked as `To verify`.

Mandatory watchlist keywords for `OD all papers`:

```text
YOLO26, YOLOv26, YOLO 26, Ultralytics YOLO26, YOLO-Master, YOLOv13, YOLOv12,
RT-DETRv4, D-FINE, DEIM, RF-DETR, RF-DETR 2, Co-DETR, Relation-DETR,
Grounding DINO, DINO-X, BiRefNet, CD-ViTO, UniFS
```

---

## 11. Final validation checklist

Before returning files, explicitly verify:

- [ ] 3D Object Detection excluded.
- [ ] Four categories only.
- [ ] Every required dataset has its own worksheet.
- [ ] Every required dataset has 10 rows whenever public results exist.
- [ ] `OD all papers` includes 2026 latest methods.
- [ ] `OD all papers` is sorted by `年月` from newest to oldest.
- [ ] YOLO26 / YOLOv26 was searched using all aliases.
- [ ] Each dataset sheet is sorted by mAP / primary metric, not FPS.
- [ ] Each dataset uses its own valid metric, including SOD and FSOD metric rules.
- [ ] AP50 is not treated as COCO mAP.
- [ ] Different splits are not mixed without marking.
- [ ] RGB SOD uses S-measure / E-measure / F-measure / MAE when mAP is unavailable.
- [ ] Few-shot rows include shot settings.
- [ ] Each row has source links when available.
- [ ] GitHub links are included when available.
- [ ] `連結` and `GitHub` are clickable hyperlinks in Excel and Markdown outputs.
- [ ] Color rules are applied.
- [ ] Header row is frozen.
- [ ] Filters are enabled.
- [ ] Text wrapping is enabled.
- [ ] Column widths are consistent.
- [ ] Incomplete / uncertain results are marked `To verify`.

---

## 12. Row color rules

Apply the following row background colors based on `類別`:

| Category | Color | Hex |
|---|---:|---|
| General Object Detection-based | Light blue | #D9EAF7 |
| Real-Time Object Detection-based | Light green | #D9EAD3 |
| RGB Salient Object Detection-based | Light orange | #FCE5CD |
| Few-Shot Object Detection-based | Light purple | #EADCF8 |

Rules:

- The `類別` column must contain the category name.
- Apply color to the full row.
- Apply the same rules in `OD all papers` and all dataset sheets.
- For hybrid methods, color by the primary category and explain hybrid nature in remarks.

---

## 13. Common mistakes to avoid

1. Do not include 3D Object Detection.
2. Do not include LiDAR / BEV / point-cloud detection.
3. Do not include medical detection unless the dataset is explicitly listed.
4. Do not only search COCO.
5. Do not only use Papers with Code.
6. Do not ignore official docs or GitHub releases.
7. Do not miss arXiv / cs.CV preprints.
8. Do not miss 2025–2026 methods.
9. Do not miss YOLO26 / YOLOv26 aliases.
10. Do not rank by FPS.
11. Do not mix AP50 with COCO mAP.
12. Do not mix different splits without marking.
13. Do not hide uncertainty; mark uncertain rows as `To verify`.
14. Do not omit source links.
15. Do not omit important methods only because GitHub is unavailable.

---

## 14. Recommended final response after completing the task

Return links to:

```text
Object_Detection_Papers_Ranking_2021_2026.xlsx
Object_Detection_Papers_Ranking_2021_2026.csv
Object_Detection_Papers_Ranking_Report.md
```

Also summarize:

- number of total rows
- number of datasets
- number of rows by category
- number of 2026 methods found
- datasets with fewer than 10 strictly comparable rows
- important rows marked `To verify`
