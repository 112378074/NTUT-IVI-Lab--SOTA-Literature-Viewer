# Paper Search Workflow for Computer Vision Paper Website

本文件用於指導 Agent / Claude Cowork / Research Assistant 進行 **2021–2027 年 Computer Vision 論文搜尋、驗證、整理與 Ranking 建立**。

此文件的目標不是建立網站，而是專門負責：

```text
搜尋論文 → 驗證來源 → 整理 Dataset → 建立 Ranking → 輸出 papers.json / papers.csv
```

最後產出的資料將提供給網站使用。

---

## 1. 任務目標

請搜尋並整理 2021–2027 年 Computer Vision 相關論文、方法、模型、Dataset 與 Ranking。

主要領域包括：

1. Anomaly Detection
2. Object Detection
3. Classification
4. Anomaly Synthesis

其中：

```text
Anomaly Synthesis 目前標註為「待開發中」，暫時不進行論文搜尋與 Ranking 整理。
```

目前主要搜尋與整理以下三個領域：

1. Anomaly Detection
2. Object Detection
3. Classification

---

## 2. 核心原則

請嚴格遵守以下原則：

```text
不同 Dataset 不混排。
不同 Evaluation Protocol 不強行比較。
沒有明確分數與來源的論文不放入正式 Rank。
每一筆資料都必須可追溯。
不捏造論文、方法、分數、GitHub、Venue、DOI。
See All 頁面只顯示單一 Dataset 的完整 Ranking。
```

---

## 3. 資料使用限制

請遵守以下規則：

1. 不大量下載 PDF。
2. 不繞過 CAPTCHA。
3. 不破解登入或跳過資料庫限制。
4. 不高速爬取 IEEE Xplore、Scopus、Web of Science、ScienceDirect、Google Scholar。
5. 對 IEEE、Scopus、Web of Science、ScienceDirect、Google Scholar，只使用正常搜尋、摘要閱讀、metadata 匯出與人工驗證。
6. 優先使用開放來源。
7. 若資料需要學校授權，請在使用者已登入學校 VPN 或授權資料庫後，再進行正常搜尋與匯出。
8. 若資料無法確認，請標註 `verified = Partial` 或 `verified = No`。
9. 2027 年資料可能尚未完整，不可自行生成不存在的 2027 年論文。
10. 無法確認 Ranking 的論文可放入 Related Papers，但不可放入正式 Rank。

---

## 4. 搜尋年份範圍

搜尋範圍：

```text
2021–2027
```

年份判斷規則：

1. 若論文已有正式出版版本，使用正式出版年月。
2. 若尚未正式出版，使用 arXiv 首次發布年月。
3. 若同一篇論文同時有 arXiv 與正式會議 / 期刊版本，優先使用正式版本。
4. 若只能確認年份，`year_month` 可填 `YYYY`。
5. 若能確認年月，`year_month` 請填 `YYYY-MM`。

---

## 5. 優先搜尋來源

搜尋來源優先順序如下：

### 5.1 第一優先：Ranking / Dataset / Benchmark 來源

1. Papers with Code
2. 官方 Dataset / Benchmark 網站
3. 官方 Leaderboard
4. GitHub README benchmark table

用途：

- 確認 Dataset
- 確認 Task
- 確認常用 Metrics
- 確認 Leaderboard / Ranking
- 找出 Best Model 與 Top models

---

### 5.2 第二優先：開放論文來源

1. arXiv
2. CVF Open Access
3. OpenReview
4. Semantic Scholar
5. OpenAlex
6. Crossref

用途：

- 搜尋 2021–2027 年最新方法
- 確認論文標題、作者、年份、Venue
- 補 DOI 或 arXiv ID
- 找正式版本

---

### 5.3 第三優先：程式碼來源

1. GitHub
2. Papers with Code code link
3. 官方 Project Page

用途：

- 確認是否有官方 GitHub
- 確認 repo 是否與論文一致
- 確認 benchmark 分數是否出現在 README
- 區分 official / unofficial implementation

---

### 5.4 第四優先：授權資料庫與人工驗證

1. IEEE Xplore
2. Scopus
3. Web of Science
4. ScienceDirect
5. Google Scholar

用途：

- 確認正式出版資訊
- 確認期刊 / 會議名稱
- 確認 DOI
- 查缺補漏
- 不進行大量自動爬取

---

## 6. 主領域與子領域

### 6.1 Anomaly Detection

需整理以下子領域：

1. Anomaly Detection
2. Unsupervised Anomaly Detection
3. One-Class Anomaly Detection
4. Supervised Anomaly Detection
5. Graph Anomaly Detection

常見 Dataset 範例：

- MVTec AD
- VisA
- MPDD
- BTAD
- Real-IAD
- MVTec LOCO AD
- DAGM
- MVTec 3D-AD
- CIFAR-10 / CIFAR-100 for anomaly detection
- MNIST / Fashion-MNIST for anomaly detection
- Graph anomaly datasets such as Cora, Citeseer, PubMed, Reddit, Amazon, Yelp, Elliptic, T-Finance

實際 Dataset 請以 Papers with Code、官方 benchmark、論文實驗設定與 leaderboard 為準。

---

### 6.2 Object Detection

需整理以下子領域：

1. Object Detection
2. 3D Object Detection
3. Real-time Object Detection
4. RGB Salient Object Detection
5. Few-Shot Object Detection

常見 Dataset 範例：

#### Object Detection

- COCO
- PASCAL VOC
- Objects365
- Open Images
- CrowdHuman
- LVIS
- DOTA
- VisDrone

#### 3D Object Detection

- KITTI
- Waymo Open Dataset
- nuScenes
- SUN RGB-D
- ScanNet
- ONCE
- Argoverse

#### Real-time Object Detection

- COCO
- PASCAL VOC
- Objects365
- VisDrone

#### RGB Salient Object Detection

- DUTS
- DUT-OMRON
- ECSSD
- HKU-IS
- PASCAL-S
- SOD
- SOC

#### Few-Shot Object Detection

- PASCAL VOC Few-Shot
- COCO Few-Shot
- LVIS Few-Shot
- FSOD

實際 Dataset 請以 Papers with Code、官方 benchmark、論文實驗設定與 leaderboard 為準。

---

### 6.3 Classification

網站中可顯示為 Image Classification。

需整理以下子領域：

1. Image Classification
2. Few-Shot Image Classification
3. Fine-Grained Image Classification
4. Semi-Supervised Image Classification
5. Small Data Image Classification

常見 Dataset 範例：

#### Image Classification

- ImageNet
- CIFAR-10
- CIFAR-100
- Tiny ImageNet
- Places365

#### Few-Shot Image Classification

- miniImageNet
- tieredImageNet
- CUB
- CIFAR-FS
- FC100
- Omniglot

#### Fine-Grained Image Classification

- CUB-200-2011
- Stanford Cars
- FGVC-Aircraft
- NABirds
- iNaturalist
- Food-101
- Oxford Flowers

#### Semi-Supervised Image Classification

- CIFAR-10
- CIFAR-100
- ImageNet
- STL-10
- SVHN

#### Small Data Image Classification

- CIFAR-10
- CIFAR-100
- Tiny ImageNet
- Oxford Flowers
- Caltech-101
- Caltech-256
- CUB-200-2011

實際 Dataset 請以 Papers with Code、官方 benchmark、論文實驗設定與 leaderboard 為準。

---

## 7. 搜尋流程總覽

每一個子領域請依照以下流程進行：

```text
Step 1：確認子領域
Step 2：搜尋該子領域常用 Dataset
Step 3：針對每個 Dataset 找 Leaderboard / Ranking
Step 4：搜尋 2021–2027 年方法與論文
Step 5：確認每篇論文的 Metrics、Paper Link、GitHub Link、Venue、Status
Step 6：建立單一 Dataset 的 Ranking
Step 7：輸出 papers.json / papers.csv
Step 8：更新 datasets.json 中的 Best Model
```

---

## 8. 詳細搜尋流程

### Step 1：確認子領域

先選定一個子領域，例如：

```text
Unsupervised Anomaly Detection
```

請記錄：

- domain
- subdomain
- 可能使用的 datasets
- 可能使用的 metrics
- 代表性方法

---

### Step 2：搜尋該子領域常用 Dataset

優先查詢：

1. Papers with Code task page
2. Papers with Code dataset page
3. 官方 benchmark page
4. Survey paper
5. Recent SOTA paper experiment section

搜尋語法範例：

```text
"unsupervised anomaly detection datasets"
"Papers with Code unsupervised anomaly detection datasets"
"MVTec AD leaderboard anomaly detection"
"industrial anomaly detection benchmark 2021 2022 2023 2024 2025"
```

請輸出 Dataset 清單，並為每個 Dataset 記錄：

- dataset_name
- domain
- subdomain
- common_metrics
- official_website
- leaderboard_url
- notes

---

### Step 3：針對每個 Dataset 找 Ranking

每一個 Dataset 都要單獨建立 Ranking。

搜尋語法範例：

```text
"MVTec AD anomaly detection leaderboard"
"MVTec AD Image AUROC Pixel AUROC 2021 2025"
"MVTec AD UniNet GLASS EfficientAD PatchCore"
"COCO object detection leaderboard AP 2021 2025"
"miniImageNet few-shot classification 5-way 1-shot leaderboard"
```

請注意：

1. Dataset Ranking 只能包含同一 Dataset。
2. 不同 protocol 不要強行混排。
3. 若分數來自不同來源，要在 notes 註明。
4. 若缺乏可比性，放入 Related Papers，不放入正式 Rank。

---

### Step 4：搜尋 2021–2027 年方法與論文

針對每個 Dataset 與子領域，搜尋 2021–2027 年的新方法。

搜尋語法範例：

```text
site:arxiv.org "MVTec AD" "anomaly detection" "2024"
site:openaccess.thecvf.com "MVTec AD" "anomaly detection"
site:github.com "MVTec AD" "anomaly detection"
"COCO object detection" "CVPR 2024" "GitHub"
"few-shot image classification miniImageNet 2023"
"RGB salient object detection DUTS S-measure 2024"
```

每篇候選論文需記錄：

- paper_title
- authors
- year_month
- venue
- status
- paper_url
- github_url
- dataset
- metrics
- source
- notes

---

### Step 5：確認 Paper Link

Paper Link 優先順序：

1. 官方 conference / journal page
2. CVF Open Access
3. arXiv
4. DOI page
5. Publisher page
6. Project page

若同一篇論文有多個版本：

```text
正式會議 / 期刊版本 > CVF / Publisher > arXiv > Project Page
```

---

### Step 6：確認 GitHub Link

GitHub Link 優先順序：

1. 論文官方 GitHub
2. 作者個人頁或 project page 連出的 GitHub
3. Papers with Code 連出的 GitHub
4. Unofficial implementation

若是非官方實作，請在 notes 標註：

```text
unofficial implementation
```

若沒有 GitHub，填入：

```text
N/A
```

---

### Step 7：確認 Venue 與 Status

Status 請填以下三種之一：

```text
Conference
Journal
Preprint
```

判斷方式：

- CVPR / ICCV / ECCV / NeurIPS / ICLR / AAAI / ACM MM / WACV 等：Conference
- IEEE Transactions / Pattern Recognition / Information Fusion / Expert Systems with Applications 等：Journal
- arXiv only：Preprint

若尚未確認正式收錄，先填：

```text
Preprint
```

並在 verified 欄位填：

```text
Partial
```

---

### Step 8：確認 Metrics

每筆 Ranking 資料必須填寫實際 Metrics。

Metrics 格式範例：

```json
{
  "Image AUROC": 99.9,
  "Pixel AUROC": 99.3,
  "Pixel AP": 70.0,
  "PRO": 96.0
}
```

或：

```json
{
  "AP": 56.8,
  "AP50": 74.2,
  "AP75": 61.5
}
```

或：

```json
{
  "Top-1 Acc": 88.5,
  "Top-5 Acc": 98.7
}
```

若只有部分指標，照實填寫，不要補假資料。

---

### Step 9：建立 Ranking

Ranking 原則：

1. 每個 Dataset 獨立排序。
2. 排序依據該 Dataset / Task 的主要指標。
3. 排序指標需寫入 `dataset_common_metrics`。
4. 若論文使用不同輸入尺寸、backbone、pretraining、test protocol，需在 notes 標註。
5. 若不可公平比較，不放入正式 Rank。

範例：

```text
MVTec AD Ranking
Rank 1: UniNet
Rank 2: GLASS
Rank 3: EfficientAD
...
```

不可寫成：

```text
Anomaly Detection Ranking
Rank 1: UniNet on MVTec AD
Rank 2: Method A on VisA
Rank 3: Method B on BTAD
```

因為這是不同 Dataset 混排。

---

## 9. 各領域常用 Metrics

### 9.1 Anomaly Detection

常用指標：

- Image AUROC
- Pixel AUROC
- Pixel AP
- PRO
- AUPRO
- F1-score
- AUPRC

排序建議：

1. 若是 image-level anomaly detection，以 Image AUROC 為主要排序。
2. 若是 pixel-level anomaly segmentation，以 Pixel AUROC / Pixel AP / PRO 為主要排序。
3. 若論文同時報告四項指標，可保留全部指標。
4. 若不同方法只報告不同指標，不可強行混排。

---

### 9.2 Object Detection

常用指標：

- AP
- AP50
- AP75
- APS
- APM
- APL
- mAP

排序建議：

1. COCO 以 AP / mAP 為主。
2. PASCAL VOC 可使用 mAP@0.5。
3. DOTA 使用 mAP。
4. CrowdHuman 可包含 MR、AP、Recall。
5. 若 protocol 不同，需標註。

---

### 9.3 Real-time Object Detection

常用指標：

- AP
- AP50
- FPS
- Latency
- Params
- FLOPs

排序建議：

1. 不同硬體上的 FPS 不可直接混排。
2. 不同輸入尺寸不可直接混排。
3. Ranking 可優先以 AP 排序，並附上 FPS / Params / FLOPs。
4. 若要做 real-time 綜合排名，需明確標註硬體與測試條件。

---

### 9.4 3D Object Detection

常用指標：

- mAP
- NDS
- AP3D
- BEV AP
- mAPH

排序建議：

1. nuScenes 常用 NDS / mAP。
2. KITTI 常用 AP3D / BEV AP。
3. Waymo 常用 mAP / mAPH。
4. 不同 difficulty level 不可混排。

---

### 9.5 RGB Salient Object Detection

常用指標：

- S-measure
- F-measure
- E-measure
- MAE

排序建議：

1. S-measure、F-measure、E-measure 越高越好。
2. MAE 越低越好。
3. 排名時需明確說明採用哪一個主指標。

---

### 9.6 Few-Shot Object Detection

常用指標：

- nAP
- nAP50
- nAP75
- bAP
- mAP
- 1-shot / 2-shot / 3-shot / 5-shot / 10-shot

排序建議：

1. shot 數不同不可混排。
2. base / novel split 不同不可混排。
3. VOC split 1 / 2 / 3 不可直接混排，需分開或標註。

---

### 9.7 Image Classification

常用指標：

- Top-1 Accuracy
- Top-5 Accuracy
- Error Rate

排序建議：

1. ImageNet 以 Top-1 Accuracy 為主要排序。
2. CIFAR-10 / CIFAR-100 通常以 Accuracy 排序。
3. 不同 training data 或 extra data 設定需標註。

---

### 9.8 Few-Shot Image Classification

常用指標：

- 5-way 1-shot Accuracy
- 5-way 5-shot Accuracy
- 95% confidence interval

排序建議：

1. 1-shot 與 5-shot 不可混排。
2. transductive / inductive setting 不可混排。
3. backbone 不同可記錄，但不要隱藏差異。

---

### 9.9 Fine-Grained Image Classification

常用指標：

- Top-1 Accuracy
- Classification Accuracy

排序建議：

1. 以 Top-1 Accuracy 為主。
2. 是否使用 bounding box annotation / part annotation 需註明。
3. 是否使用 extra training data 需註明。

---

### 9.10 Semi-Supervised Image Classification

常用指標：

- Accuracy
- Error Rate

排序建議：

1. label 數不同不可混排。
2. CIFAR-10 40 labels、250 labels、4000 labels 要分開。
3. ImageNet 1% labels、10% labels 要分開。
4. 若使用 extra data，需註明。

---

## 10. 正式 Rank vs Related Papers

### 10.1 可放入正式 Rank 的條件

一篇論文 / 方法可放入正式 Rank，必須符合：

1. 有明確 Dataset。
2. 有明確 Metrics。
3. 有可追溯來源。
4. 使用相同或可比較的 evaluation protocol。
5. 可確認 Paper Link。
6. 可確認 Year / Venue / Status。
7. 若有 GitHub，需確認是否官方。

---

### 10.2 只能放入 Related Papers 的情況

以下情況不可放入正式 Rank：

1. 沒有報告明確分數。
2. 只提出方法，沒有在該 Dataset 上評估。
3. Dataset 不一致。
4. Evaluation protocol 不一致且無法公平比較。
5. 只在私有資料集測試。
6. 分數來源無法追溯。
7. GitHub README 分數與論文不一致，且無法確認。
8. 只有二手引用，找不到原始來源。

---

## 11. 資料欄位格式

請輸出 `papers.json` 與 `papers.csv`，欄位需一致。

### 11.1 papers.json 欄位

每筆資料至少包含：

```json
{
  "domain": "",
  "subdomain": "",
  "dataset": "",
  "rank": "",
  "model_name": "",
  "category": "",
  "metrics": {},
  "dataset_common_metrics": [],
  "ranking_metric": "",
  "paper_title": "",
  "authors": "",
  "year_month": "",
  "venue": "",
  "status": "",
  "paper_url": "",
  "github_url": "",
  "github_type": "official / unofficial / N/A",
  "doi": "",
  "code_available": "Yes / No",
  "innovation_summary": "",
  "source": "",
  "verified": "Yes / Partial / No",
  "notes": ""
}
```

### 11.2 papers.csv 欄位

CSV 欄位順序建議：

```text
domain,
subdomain,
dataset,
rank,
model_name,
category,
metrics,
dataset_common_metrics,
ranking_metric,
paper_title,
authors,
year_month,
venue,
status,
paper_url,
github_url,
github_type,
doi,
code_available,
innovation_summary,
source,
verified,
notes
```

---

## 12. Dataset 資料欄位格式

請輸出 `datasets.json`。

每筆 Dataset 至少包含：

```json
{
  "dataset_id": "",
  "dataset_name": "",
  "domain": "",
  "subdomain": "",
  "best_model": "",
  "best_model_paper_url": "",
  "best_model_github_url": "",
  "common_metrics": [],
  "ranking_metric": "",
  "leaderboard_url": "",
  "official_website": "",
  "status": "active / under_review / insufficient_data",
  "notes": ""
}
```

注意：

```text
datasets.json 中的 best_model 必須與 papers.json 中該 dataset 的 Rank 1 一致。
```

---

## 13. Subdomain 資料欄位格式

請輸出 `subdomains.json`。

格式範例：

```json
[
  {
    "domain": "Anomaly Detection",
    "subdomains": [
      "Anomaly Detection",
      "Unsupervised Anomaly Detection",
      "One-Class Anomaly Detection",
      "Supervised Anomaly Detection",
      "Graph Anomaly Detection"
    ]
  },
  {
    "domain": "Object Detection",
    "subdomains": [
      "Object Detection",
      "3D Object Detection",
      "Real-time Object Detection",
      "RGB Salient Object Detection",
      "Few-Shot Object Detection"
    ]
  },
  {
    "domain": "Classification",
    "subdomains": [
      "Image Classification",
      "Few-Shot Image Classification",
      "Fine-Grained Image Classification",
      "Semi-Supervised Image Classification",
      "Small Data Image Classification"
    ]
  },
  {
    "domain": "Anomaly Synthesis",
    "subdomains": [],
    "status": "under_development"
  }
]
```

---

## 14. 建議輸出資料夾

```text
paper-website/
│
├── data/
│   ├── papers.json
│   ├── papers.csv
│   ├── datasets.json
│   └── subdomains.json
│
├── sources/
│   ├── arxiv_results.csv
│   ├── cvf_results.csv
│   ├── github_results.csv
│   ├── paperswithcode_results.csv
│   ├── openalex_results.csv
│   ├── crossref_results.csv
│   ├── semantic_scholar_results.csv
│   ├── ieee_export.csv
│   ├── scopus_export.csv
│   ├── wos_export.csv
│   └── sciencedirect_export.csv
│
└── logs/
    ├── search_log.md
    ├── verification_log.md
    └── unresolved_items.md
```

---

## 15. 搜尋紀錄要求

請建立 `logs/search_log.md`，記錄每次搜尋：

```markdown
## Search Record

- Date:
- Domain:
- Subdomain:
- Dataset:
- Query:
- Source:
- Useful Results:
- Excluded Results:
- Notes:
```

---

## 16. 驗證紀錄要求

請建立 `logs/verification_log.md`，記錄每筆資料的驗證狀態：

```markdown
## Verification Record

- Paper Title:
- Model:
- Dataset:
- Metrics:
- Paper URL:
- GitHub URL:
- Source:
- Verified: Yes / Partial / No
- Issues:
- Notes:
```

---

## 17. 無法解決項目紀錄

請建立 `logs/unresolved_items.md`，記錄無法確認的資料：

```markdown
## Unresolved Item

- Domain:
- Subdomain:
- Dataset:
- Paper / Model:
- Missing Information:
- Reason:
- Suggested Next Step:
```

---

## 18. 每個 Dataset 的最低整理要求

每個 Dataset 盡量整理 Top 10 models。

若不足 Top 10，請標註：

```text
available records only
```

每個 Dataset 至少要有：

1. Dataset name
2. Subdomain
3. Common metrics
4. Ranking metric
5. Rank table
6. Best Model
7. Paper links
8. GitHub links
9. Source
10. Notes

---

## 19. 搜尋 Query 範本

### 19.1 Anomaly Detection

```text
"Papers with Code" "MVTec AD" "anomaly detection"
"MVTec AD" "Image AUROC" "Pixel AUROC" "2024"
"MVTec AD" "anomaly detection" "CVPR 2024"
"industrial anomaly detection" "MVTec AD" "GitHub"
"VisA anomaly detection benchmark"
"MPDD anomaly detection leaderboard"
"Graph anomaly detection benchmark AUROC 2024"
```

---

### 19.2 Object Detection

```text
"Papers with Code" "COCO object detection leaderboard"
"COCO object detection AP 2024 GitHub"
"real-time object detection COCO FPS AP 2024"
"YOLOv10 COCO AP GitHub"
"RT-DETR COCO AP GitHub"
"DOTA object detection mAP 2024"
"KITTI 3D object detection AP3D 2024"
"nuScenes 3D object detection NDS mAP 2024"
"few-shot object detection COCO nAP 2024"
```

---

### 19.3 Classification

```text
"Papers with Code" "ImageNet classification leaderboard"
"ImageNet Top-1 accuracy 2024 GitHub"
"few-shot image classification miniImageNet 5-way 1-shot 2024"
"fine-grained image classification CUB-200-2011 2024"
"semi-supervised image classification CIFAR-10 40 labels 2024"
"small data image classification benchmark 2024"
```

---

## 20. 最終交付內容

完成搜尋後，請輸出以下檔案：

```text
data/papers.json
data/papers.csv
data/datasets.json
data/subdomains.json
logs/search_log.md
logs/verification_log.md
logs/unresolved_items.md
```

其中：

- `papers.json`：網站主要 Ranking 資料來源。
- `papers.csv`：方便人工檢查與 Excel 編輯。
- `datasets.json`：Dataset card 與 Best Model 資料來源。
- `subdomains.json`：網站主領域與子領域架構。
- `search_log.md`：搜尋過程紀錄。
- `verification_log.md`：資料驗證紀錄。
- `unresolved_items.md`：尚未確認或需人工處理的項目。

---

## 21. 最終檢查清單

完成資料整理前，請確認：

- [ ] 每筆資料都有 domain。
- [ ] 每筆資料都有 subdomain。
- [ ] 每筆資料都有 dataset。
- [ ] 每個 dataset 都是獨立 ranking。
- [ ] 沒有不同 dataset 混排。
- [ ] 沒有不同 protocol 強行混排。
- [ ] 每筆正式 Rank 都有明確分數。
- [ ] 每筆正式 Rank 都有 paper link。
- [ ] 每筆正式 Rank 都有 source。
- [ ] GitHub link 已確認 official / unofficial / N/A。
- [ ] Venue 與 Status 已確認。
- [ ] 2027 年資料沒有捏造。
- [ ] Best Model 與 Rank 1 一致。
- [ ] 無法確認的資料已放入 unresolved_items.md。
- [ ] 重要但不能排名的論文已放入 Related Papers 或 notes。

---

## 22. 給 Agent 的執行指令

請依照以下順序執行：

```text
1. 不要先製作網站。
2. 先依照本文件搜尋與整理論文資料。
3. 先建立 subdomains.json。
4. 再建立 datasets.json。
5. 接著針對每個 dataset 建立 papers.json / papers.csv。
6. 每個 dataset 的 See All ranking 只包含該 dataset。
7. 若資料不足 Top 10，標註 available records only。
8. 每筆資料都要有來源與驗證狀態。
9. 最後再將整理好的資料交給網站模板使用。
```

---

## 23. 最重要的規則

```text
不要為了湊滿 Top 10 而編造資料。
不要把不同 Dataset 的結果混在一起。
不要把不同 evaluation setting 的結果強行排序。
不要把沒有分數的論文放入正式 Ranking。
所有 Ranking 都必須能回到原始來源驗證。
```
