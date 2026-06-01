---
name: cv-paper-pipeline
summary: "Automatically search, classify, verify, rebuild, publish, and notify for the CV SOTA Literature Viewer."
description: >
  Search public academic sources (arXiv and CVF Open Access) for Computer Vision
  papers in Anomaly Detection / Object Detection / Image Classification, classify
  each new paper by task, dataset split, protocol, metric, and source quality,
  then update the website data directly without using Excel as an intermediate
  source of truth. The pipeline rebuilds the literature-viewer website
  (index.html and related JSON/data blocks), pushes to GitHub Pages, and emails
  a summary to configured recipients after each scheduled run. Use this skill
  when the user wants to update the paper database, search new CV papers, rebuild
  the website, change classification/verification rules, or adjust schedule/email
  settings.
---

# CV Paper Pipeline

End-to-end workflow: **search public sources → classify & verify → website data → GitHub Pages → email**.

```text
arXiv / CVF Open Access
        │
        ▼
New-paper candidates
        │
        ▼
Classify & verify gate
(task → dataset split → protocol → metric → source quality)
        │
        ▼
Website data directly
(*_data.json / embedded index.html data blocks)
        │
        ▼
index.html → git push → GitHub Pages
        │
        ▼
Email summary
```

## Core principle

The website data is updated directly. Do **not** use Excel workbooks as the update target or single source of truth.

All new papers must pass classification and source verification before entering any sortable leaderboard. A paper can be shown as a catalog or protocol-specific entry before metric verification, but it must not be ranked with a numeric score unless the dataset, split, protocol, metric, and source are confirmed.

## Project layout

| File / path | Role |
| --- | --- |
| `index.html` | Single-file website; data is injected or embedded for the viewer. |
| `*_data.json` | Generated website data for AD / OD / CLS / other enabled domains. |
| `cls_supplemental.json` | Optional hand-curated CLS supplemental data, if still used by the site. |
| `scripts/update_papers.py` | Main pipeline — arXiv + CVF search, classification staging, website-data regeneration, inject/rebuild, push, email. |
| `scripts/run_update.bat` | Windows Task Scheduler wrapper. |
| `scripts/.env` | SMTP credentials and notification recipients; gitignored. |
| `.claude/skills/cv-paper-pipeline/scripts/fetch_cvf.py` | Standalone CVF scraper for ad-hoc conference scans. |
| `references/verification_workflow.md` | Classification and verification rules for AD / OD / CLS papers. |

## Step 1 — Search public sources

### arXiv

arXiv search is wired into `scripts/update_papers.py`.

```bash
python scripts\update_papers.py --dry
python scripts\update_papers.py
```

A dry run previews candidates without publishing. A full run searches, stages candidates, rebuilds the website data, pushes changes, and emails the summary.

### CVF Open Access

CVF Open Access is public HTML. The scheduled pipeline scans it automatically through `process_cvf()` in `scripts/update_papers.py`.

For an ad-hoc scan of a specific conference:

```bash
python .claude\skills\cv-paper-pipeline\scripts\fetch_cvf.py --conf CVPR --year 2025
python .claude\skills\cv-paper-pipeline\scripts\fetch_cvf.py --conf ICCV --year 2025 --domain od
```

The output should be treated as **candidate metadata** only. Do not insert metric values into a leaderboard from scraped metadata alone.

## Removed sources

Do not include IEEE Xplore, Web of Science, ScienceDirect, or Scopus in this skill workflow. Do not request manual exports from those platforms as part of the scheduled pipeline.

The supported automated sources are:

- arXiv
- CVF Open Access

## Step 1c — Journal-sourced harvest (agent-driven)

In addition to the automated arXiv + CVF scan, run a **journal-targeted harvest**: pick target journals, search for AD/OD papers they published, and verify each metric against the paper's own table before it is ranked. This is agent-driven (web search + per-paper verification); the headless `.bat` cannot do it.

**Workflow (verify → add → classify → sort → never fabricate):**

1. **Pick target journals.** Use the journal list produced by the `journal-finder` skill (`期刊清單_*.xlsx`). IVAD-relevant journals by domain:
   - *Image / CV*: IEEE TPAMI, IJCV, IEEE TIP, Pattern Recognition, IEEE TCSVT, CVIU, IVC.
   - *Smart manufacturing*: IEEE TII, Computers in Industry, IEEE TIM, J. Manufacturing Systems, J. Intelligent Manufacturing, RCIM, IEEE T-ASE, IEEE TIE.
   - *AI / ML*: Information Fusion, IEEE TNNLS, Knowledge-Based Systems, Information Sciences, Neural Networks, IEEE TCYB, AIJ, Nature Machine Intelligence.
2. **Search** `WebSearch` per `journal + dataset + metric + year(2020-2026)`, e.g. `anomaly detection MVTec AD VisA "IEEE Transactions on Industrial Informatics" 2024 2025 arxiv github`, or `multispectral object detection "Information Fusion" FLIR LLVIP mAP github`.
3. **Verify the metric from an authoritative table** before ranking — read the paper's own results table from the **official GitHub repo README** or the **arXiv HTML** (`https://arxiv.org/html/<id>`). Cross-paper tables are acceptable only with an explicit note. If no accessible table exists (paywalled, no arXiv, repo has no table) → **do not rank it**; stage catalog/⏳ Pending with blank metric (e.g. IGAF/Information Fusion had no table → not added).
4. **Classify** (method-based category) and pick the **dataset + split + protocol + primary metric** per `references/verification_workflow.md`. Mind dataset-specific metrics — e.g. **MVTec AD 2 uses SegF1 / ClassF1 / AU-PRO@0.05, not AUROC**; RGB-T detection uses per-dataset mAP50 (FLIR vs M3FD vs LLVIP are different comparable groups — never cross-rank).
5. **Insert** into the correct workbook sheet (AD: `AnomalyDetection_*.xlsx`; OD: `Object_Detection_*.xlsx`) with the value in the right column, a `✅ … verified` source note citing repo/arXiv table + dataset + metric, and dupe-check first. The website sorts client-side, so appending is fine.
6. **Regenerate + reinject + push + email** (Step 3 below).

**Verified examples added this way:** Hyper-YOLO-L & YOLO-MS (IEEE TPAMI → COCO val2017), CDO (IEEE TII) & MSFlow (IEEE TNNLS) → MVTec AD + VisA, COMO (Information Fusion → LLVIP). Each carries a source note pointing at the official repo/arXiv table.

**Never fabricate:** if a value isn't in an authoritative table, leave it blank with the reason — do not pad a leaderboard. Re-verify before changing any existing number (the MVTec AD 2 sheet contained placeholder values and was rebuilt from the VAND 3.0 report + RoBiS paper tables).

## Step 2 — Classify & verify before any leaderboard insert

A new AD / OD / CLS paper is never auto-inserted into a standard sortable leaderboard with a numeric score. It first enters as a candidate or catalog entry with a blank score and a pending / needs-verification status.

Only after the following checks pass can it become a ranked row.

### 2.1 Task type

Identify the actual task from the title, abstract, method, experiment section, and result tables.

For OD, do not classify a paper as Standard OD just because it says “object detection.” Possible task types include:

- Standard 2D closed-set Object Detection
- Real-Time Object Detection
- 3D / BEV Object Detection
- RGB Salient Object Detection / DIS / Co-SOD
- Few-Shot Object Detection
- Cross-Domain Few-Shot Object Detection
- Open-Vocabulary Object Detection
- Zero-Shot Object Detection
- Open-World Object Detection
- Long-Tail Object Detection
- Domain Adaptation / Domain Generalization OD
- Robustness / Security / Attack OD
- Referring / Grounded / MLLM-based Detection
- Oriented / Remote-Sensing Object Detection
- Other specialized OD task

For AD, identify whether it is:

- Standard unsupervised AD
- One-class AD
- Supervised AD
- Zero-shot AD
- Few-shot AD
- Continual AD
- 3D / multimodal AD
- Logical AD
- Segmentation-only / localization-only AD
- Synthetic anomaly generation / anomaly synthesis
- Other specialized AD protocol

For CLS, identify whether it is:

- Standard supervised image classification
- Semi-supervised classification
- Self-supervised / representation learning evaluation
- Few-shot / small-data classification
- Long-tail classification
- Robustness / OOD / domain-shift classification
- Medical / fine-grained / specialized classification

### 2.2 Dataset and split

Record the exact dataset and split used by the paper.

Examples:

- COCO test-dev
- COCO val2017 / minival / COCO 2017 val
- PASCAL VOC 2007 test
- LVIS v1.0 val
- LVIS minival
- CrowdHuman
- Waymo 2D
- CPPE-5
- OV-COCO
- OV-LVIS
- MVTec AD
- VisA
- BTAD
- MPDD
- MVTec LOCO AD
- MVTec 3D-AD
- ImageNet-1K
- ImageNet-ReaL / A / R / Sketch
- CIFAR-10 / CIFAR-100
- miniImageNet / tieredImageNet / CIFAR-FS / FC100 / CUB

Never mix val, minival, and test-dev. If the paper reports LVIS minival, do not insert it into LVIS v1.0 val unless the page explicitly allows that split and the row clearly states it.

### 2.3 Protocol

Record the actual protocol. Examples:

- Fully-supervised closed-set
- Open-vocabulary
- Zero-shot
- Few-shot
- Cross-domain few-shot
- Open-world
- Long-tail / rare-class
- Domain adaptation
- Domain generalization
- Robustness / corruption / attack
- Incremental
- Oriented bounding box
- 3D / BEV
- SOD / DIS / Co-SOD
- Referring / grounding
- Standard unsupervised AD
- Zero-shot AD
- Few-shot AD
- Continual AD
- Standard supervised CLS
- Semi-supervised CLS
- Self-supervised evaluation
- Long-tail CLS

Same dataset + different protocol is allowed on the same dataset page only when clearly labeled. Do not call verified non-standard protocol rows “unverified.”

### 2.4 Metric

Use the metric actually reported by the paper and expected by that dataset/protocol.

Standard OD examples:

- COCO: AP primary; AP50 / AP75 / APS / APM / APL secondary
- PASCAL VOC: mAP@0.5
- LVIS standard: AP
- CrowdHuman: AP primary, MR⁻² secondary
- Waymo 2D: AP/L1 primary, AP/L2 secondary
- CPPE-5: AP50 or the dataset-paper specified metric

Non-standard OD examples:

- Open-Vocabulary OD: APN50, APB50, AP_rare, APr, APm, APAll
- Few-Shot OD: novel AP, nAP, AP50 novel, 1/5/10/30-shot AP
- SOD / DIS / Co-SOD: S-measure, F-measure, E-measure, MAE
- Open-World OD: U-Recall, WI, A-OSE, mAP
- Domain Adaptation OD: target-domain AP
- Robustness / Security OD: robust AP, mAP drop, ASR
- Oriented OD: oriented mAP / AP50
- 3D / BEV OD: 3D AP, BEV AP, NDS

AD examples:

- Image AUROC / I-AUROC
- Pixel AUROC / P-AUROC
- Pixel AP
- AU-PRO / PRO
- Accuracy, when the paper genuinely uses classification-style AD

CLS examples:

- Top-1 accuracy
- Top-5 accuracy
- 5-way k-shot accuracy
- Long-tail many/medium/few-shot accuracy
- Robustness / OOD accuracy or error rate

Do not convert one metric into another. Do not infer COCO-O AP from normal COCO AP. Do not treat S-measure, novel AP, U-Recall, ASR, or classification accuracy as standard COCO AP.

### 2.5 Placement decision

Use these rules.

#### A. Standard leaderboard

Place the paper in the Standard Leaderboard only if all are true:

- same dataset
- same split
- same standard protocol
- same primary metric
- numeric score is source-verified

Status: `✅ Standard verified`.

#### B. Same dataset, different protocol

Keep the row on the same dataset page if useful, but place it under a protocol-specific group.

Examples:

- `🟣 Open-vocabulary verified`
- `🟠 Zero-shot verified`
- `🟢 Few-shot verified`
- `🟡 Non-standard protocol verified`
- `🔵 Long-tail verified`

Do not label these as unverified once the source has been checked.

#### C. Different dataset

Move it to the correct dataset page. If that dataset does not exist, create a new dataset entry under the correct domain/sub-area or keep it as a catalog entry until the page exists.

#### D. Different metric

Keep the real metric. Add it to the correct task-specific leaderboard or protocol group. Do not force it into a standard AP / Top-1 / AUROC ranking.

#### E. No authoritative source

Do not insert the score into any sortable leaderboard.

Use:

- Status: `⚠️ No authoritative source`
- Score: blank / `—`
- Placement: catalog / needs-reference / related only

#### F. Placeholder, duplicate, invalid link, or unidentified method

Remove it from public leaderboard data or keep it only in an internal needs-reference queue. Do not show it as a ranked public result.

### 2.6 Source note

Every numeric score must record a source note.

The note must include:

- paper name or arXiv ID
- table number or page when available
- dataset split
- metric name
- whether it is original-paper verified, official-leaderboard verified, dataset-paper baseline, or cross-paper verified

Examples:

- `✅ Standard verified — COCO val AP reported in Table 2 of the paper.`
- `✅ Cross-paper verified — reported by UniNet CVPR 2025 Table 1(b), not the original method paper.`
- `🟣 Open-vocabulary verified — OV-LVIS AP_rare reported in Table 3; not comparable to fully-supervised LVIS AP.`
- `⚠️ No authoritative source — value not found in original paper or benchmark table.`

No source note means the row cannot be a leaderboard row.

## Step 3 — Update the website directly

The pipeline updates the website data directly rather than appending rows to Excel workbooks.

The update process should:

1. Add new candidates to the website data structure with blank score until verified.
2. Add verified rows to the correct dataset/protocol group.
3. Regenerate `*_data.json` or the embedded data blocks used by `index.html`.
4. Re-inject or rebuild `index.html`.
5. Validate sorting and grouping.
6. Commit and push.

Example regeneration command, adapted to the current implementation:

```bash
python -c "import sys; sys.path.insert(0,'scripts'); import update_papers as up; \
ad=up.regenerate_ad_json(); od=up.regenerate_od_json(); \
cls=up.regenerate_cls_json(); asd=up.regenerate_as_json(); \
up.reinject_html(ad,od,cls,asd); print('done')"
```

Commit and push:

```bash
git add index.html *_data.json cls_supplemental.json
git commit -m "update: <what changed>"
git push
```

Do not add `.xlsx` files unless the project still keeps legacy workbooks for backup only. Website updates should not depend on editing Excel.

## Sorting and grouping rules

Default page sorting:

1. Standard comparable results first, sorted by the dataset primary metric.
2. Then protocol-specific groups, each sorted by its own primary metric.
3. Then no-source / pending rows at the bottom with no numeric ranking.

Same dataset + different protocol should stay on the dataset page when useful, but must be grouped and labeled.

Example for LVIS:

- Group 1: Fully-supervised LVIS, sorted by AP.
- Group 2: Open-vocabulary / zero-shot LVIS, sorted by AP or AP_rare.
- Group 3: Few-shot LVIS, sorted by few-shot AP.
- Group 4: Pending / no-source rows.

Example for MVTec AD:

- Group 1: Standard unsupervised AD, sorted by I-AUROC.
- Group 2: Zero-shot AD, sorted by its reported primary metric.
- Group 3: Few-shot AD, sorted by its reported primary metric.
- Group 4: 3D / multimodal / logical AD protocols, sorted within each group.

Example for ImageNet / CLS:

- Group 1: Standard supervised ImageNet-1K, sorted by Top-1.
- Group 2: Semi-supervised / self-supervised evaluations, grouped by label budget or pretraining protocol.
- Group 3: Few-shot / small-data protocols, grouped by benchmark and shot setting.
- Group 4: Robustness / OOD datasets, sorted by their own metric.

## Automation — scheduled runs

The pipeline runs unattended every Wednesday and Friday at 02:00 via Windows Task Scheduler and emails a summary to configured recipients after each run.

### What a scheduled run does

`scripts/update_papers.py` invoked by `scripts/run_update.bat` should execute:

1. arXiv fetch — recent AD / OD / CLS queries.
2. CVF Open Access scan — CVPR / ICCV / ECCV / WACV proceedings.
3. Deduplicate candidate papers.
4. Classify each new paper by domain, task type, dataset split, protocol, and metric.
5. Add candidates to website data directly with blank score and verification status unless already source-verified.
6. Never insert a numeric value into a standard ranking without source verification.
7. Regenerate `*_data.json` and re-inject / rebuild `index.html`.
8. Validate grouping and sorting.
9. `git push` to GitHub Pages.
10. Email a summary of new papers, classifications, verified rows, pending rows, and excluded/no-source rows.

### Task Scheduler entries

Two weekly tasks run as the logged-in user:

| Task name | Trigger |
| --- | --- |
| `AD-OD Paper Auto-Update Wed` | Weekly · Wednesday · 02:00 |
| `AD-OD Paper Auto-Update Fri` | Weekly · Friday · 02:00 |

Register / re-register them in an elevated shell:

```text
schtasks /Create /TN "AD-OD Paper Auto-Update Wed" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D WED /ST 02:00 /F
schtasks /Create /TN "AD-OD Paper Auto-Update Fri" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D FRI /ST 02:00 /F
```

Useful commands:

```text
schtasks /Query /TN "AD-OD Paper Auto-Update Wed"
schtasks /Run /TN "AD-OD Paper Auto-Update Wed"
```

### Email configuration

SMTP credentials live in `scripts/.env` and must stay gitignored:

```text
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=azaz31855@gmail.com
SMTP_PASSWORD=<gmail app password>
NOTIFY_TO=azaz31855@gmail.com,fctien@ntut.edu.tw
```

`NOTIFY_TO` may be comma-, semicolon-, or whitespace-separated.

If `.env` is missing or the password is empty, the run should log `email skipped` and still complete. `--no-email` skips email for a manual run.

### Caveats

- Windows scheduled tasks may run only while the user is logged in unless configured otherwise.
- Git push requires valid GitHub credentials in the environment.
- Email requires valid SMTP credentials.
- If legacy Excel files still exist, they should be treated as archival backups, not as the required update path.

## Per-domain ranking rules

### AD

Rank only within the same dataset + protocol + metric.

Common metrics:

- Image AUROC / I-AUROC
- Pixel AUROC / P-AUROC
- Pixel AP
- AU-PRO / PRO

Do not mix standard unsupervised AD with zero-shot, few-shot, continual, 3D/multimodal, logical, or segmentation-only protocols without protocol grouping.

### OD

Each dataset and protocol has its own metric.

Examples:

- COCO AP
- VOC mAP@0.5
- LVIS AP
- CrowdHuman AP / MR⁻²
- S-measure for SOD
- novel AP for FSOD
- AP_rare / APN50 for OVD
- oriented mAP for remote-sensing OD
- NDS / 3D AP for 3D/BEV OD

Never mix different protocols in one ungrouped ranking.

### CLS

Standard CLS ranking uses Top-1 as primary and Top-5 as secondary, but only within the same dataset and protocol.

Do not mix:

- standard supervised Top-1
- semi-supervised label budgets
- self-supervised linear probing / fine-tuning settings
- few-shot classification episodes
- long-tail protocols
- robustness / OOD datasets

### AS

If anomaly synthesis is disabled on the public site, do not surface it unless the user explicitly asks to enable it.

## Final report requirements

After each run, report:

1. New papers found.
2. Domain classification: AD / OD / CLS.
3. Task type / protocol classification.
4. Dataset and split.
5. Metric identified.
6. Whether the paper entered a Standard Leaderboard, protocol-specific group, catalog-only state, or no-source queue.
7. Rows added / moved / removed.
8. Verified numeric scores and their source notes.
9. Pending visual table reads.
10. Website files regenerated.
11. Git commit hash.
12. Email status.

## References

- `references/verification_workflow.md` — classify-and-verify gate for every new AD / OD / CLS paper.
- `references/source_access.md` — legacy reference only; do not use IEEE / WoS / ScienceDirect / Scopus manual exports in the current pipeline.
