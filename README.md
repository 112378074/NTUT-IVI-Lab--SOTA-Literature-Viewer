# Computer Vision Paper Website

本專案用於建立一個 **2021–2027 年 Computer Vision 論文與方法整理網站**，以階層式方式呈現不同研究領域、子領域、Dataset 與 Dataset ranking。

網站核心邏輯為：

```text
Domain
  → Subdomain
    → Dataset List
      → Dataset Card
        → See All
          → This Dataset Ranking Only
```

也就是說，使用者點選某個 Dataset 的 **See All** 後，只會看到該 Dataset 的完整排名，不會混入其他 Dataset，也不會顯示整個子領域的所有方法。

---

## 1. 專案目標

本網站整理下列主領域的 2021–2027 年論文、方法、模型、Dataset 與 Ranking：

1. Anomaly Detection
2. Object Detection
3. Classification
4. Anomaly Synthesis

其中 **Anomaly Synthesis** 目前設定為：

```text
Anomaly Synthesis：待開發中
```

先不進行資料蒐集與排名整理。

---

## 2. 主領域與子領域

### 2.1 Anomaly Detection

包含以下子領域：

1. Anomaly Detection
2. Unsupervised Anomaly Detection
3. One-Class Anomaly Detection
4. Supervised Anomaly Detection
5. Graph Anomaly Detection

### 2.2 Object Detection

包含以下子領域：

1. Object Detection
2. 3D Object Detection
3. Real-time Object Detection
4. RGB Salient Object Detection
5. Few-Shot Object Detection

### 2.3 Classification

Classification 在網站中以 **Image Classification** 為主要名稱呈現，包含以下子領域：

1. Image Classification
2. Few-Shot Image Classification
3. Fine-Grained Image Classification
4. Semi-Supervised Image Classification
5. Small Data Image Classification

### 2.4 Anomaly Synthesis

目前狀態：

```text
Under Development / 待開發中
```

---

## 3. 頁面架構

### 3.1 首頁

首頁顯示四個主領域按鈕：

- Anomaly Detection
- Object Detection
- Classification
- Anomaly Synthesis：待開發中

### 3.2 子領域頁面

點選主領域後，進入該領域的子領域列表。

例如：

```text
Object Detection
  → Object Detection
  → 3D Object Detection
  → Real-time Object Detection
  → RGB Salient Object Detection
  → Few-Shot Object Detection
```

### 3.3 Dataset 清單頁面

點選子領域後，進入該子領域的 Dataset 清單。

每個 Dataset card 僅顯示：

1. Dataset 名稱
2. Dataset 所屬子領域
3. Best Model
4. Paper Link
5. GitHub Link
6. See All 按鈕

Dataset card 不顯示：

- Best Model 的主要分數
- 常用評估指標
- 年月

完整指標、排序依據、年份、會議與來源都放在 See All 之後的完整 Ranking 頁面。

### 3.4 See All Ranking 頁面

每個 Dataset card 的 **See All** 按鈕只會進入「該 Dataset 的完整排名頁面」。

例如：

```text
Anomaly Detection
  → Unsupervised Anomaly Detection
    → MVTec AD
      → See All
        → 只顯示 MVTec AD 的完整 ranking
```

See All 頁面不能混入其他 Dataset，也不能變成整個子領域的總表。

---

## 4. Ranking 表格欄位

See All 頁面中的完整 Ranking 表格必須包含以下欄位：

| 欄位 | 說明 |
|---|---|
| Rank | 該模型在此 Dataset 的排名 |
| Model | 方法或模型名稱 |
| Category | 方法所屬類型或任務類型 |
| Dataset | 此 ranking 所屬 Dataset |
| Metrics | 該論文或方法實際回報的數值結果 |
| Dataset 常用評估指標 | 該 Dataset 或任務常用排序指標 |
| Paper Title | 論文完整標題 |
| Paper Link | 論文連結 |
| GitHub Link | 官方 GitHub 優先，無則填 N/A |
| Year/Month | 格式建議為 YYYY-MM |
| Venue | CVPR、ICCV、ECCV、NeurIPS、IEEE journal、arXiv 等 |
| Status | Conference / Journal / Preprint |
| Source | Papers with Code、arXiv、CVF、IEEE、Scopus、WoS、ScienceDirect、GitHub 等 |

---

## 5. 資料來源

優先資料來源：

1. Papers with Code
2. arXiv
3. CVF Open Access
4. GitHub
5. OpenAlex
6. Crossref
7. Semantic Scholar
8. IEEE Xplore
9. Scopus
10. Web of Science
11. ScienceDirect
12. Google Scholar

注意：

- 不大量下載 PDF。
- 不繞過 CAPTCHA。
- 不破解登入或跳過資料庫限制。
- IEEE、Scopus、Web of Science、ScienceDirect、Google Scholar 只作為正常搜尋、摘要閱讀、metadata 匯出與人工驗證來源。
- 每筆資料都必須可以追溯來源。
- 不要捏造不存在的論文、方法、分數、GitHub 連結或 DOI。

---

## 6. Ranking 規則

每個 Dataset 都要有自己的獨立 Ranking。

常見排序指標可依任務區分：

| 任務 | 常用指標 |
|---|---|
| Object Detection | mAP、AP、AP50、AP75 |
| Real-time Object Detection | mAP、FPS、Latency、Params、FLOPs |
| 3D Object Detection | mAP、NDS、AP3D、BEV AP |
| Anomaly Detection | Image AUROC、Pixel AUROC、Pixel AP、PRO |
| Graph Anomaly Detection | AUROC、AUPRC、F1-score、Precision、Recall |
| Image Classification | Top-1 Accuracy、Top-5 Accuracy |
| Few-Shot Image Classification | 1-shot Accuracy、5-shot Accuracy |
| RGB Salient Object Detection | S-measure、F-measure、E-measure、MAE |

不同 Dataset 不可混排。

不同 evaluation protocol 不可強行混排，需在 `notes` 欄位中說明。

沒有分數或沒有可確認排名的論文，不應放入正式 Rank，可另外放入 Related Papers。

---

## 7. 資料格式

### 7.1 papers.json

每筆資料建議包含：

```json
{
  "domain": "Anomaly Detection",
  "subdomain": "Unsupervised Anomaly Detection",
  "dataset": "MVTec AD",
  "rank": 1,
  "model_name": "UniNet",
  "category": "Unsupervised AD",
  "metrics": {
    "Image AUROC": "99.9",
    "Pixel AUROC": "98.8"
  },
  "dataset_common_metrics": [
    "Image AUROC",
    "Pixel AUROC",
    "Pixel AP",
    "PRO"
  ],
  "paper_title": "Paper title here",
  "authors": "Author list",
  "year_month": "2025-06",
  "venue": "CVPR",
  "status": "Conference",
  "paper_url": "https://example.com",
  "github_url": "https://github.com/example/repo",
  "doi": "N/A",
  "code_available": "Yes",
  "innovation_summary": "Brief summary of the method innovation.",
  "source": "Papers with Code",
  "verified": "Partial",
  "notes": "Official implementation; evaluation protocol should be checked."
}
```

### 7.2 datasets.json

建議格式：

```json
{
  "dataset_id": "mvtec_ad",
  "dataset_name": "MVTec AD",
  "domain": "Anomaly Detection",
  "subdomain": "Unsupervised Anomaly Detection",
  "best_model": "UniNet",
  "paper_url": "https://example.com",
  "github_url": "https://github.com/example/repo",
  "status": "active"
}
```

### 7.3 subdomains.json

建議格式：

```json
{
  "domain": "Object Detection",
  "subdomains": [
    "Object Detection",
    "3D Object Detection",
    "Real-time Object Detection",
    "RGB Salient Object Detection",
    "Few-Shot Object Detection"
  ]
}
```

---

## 8. 建議資料夾結構

```text
paper-website/
│
├── index.html
├── style.css
├── script.js
├── README.md
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
│   ├── ieee_export.csv
│   ├── scopus_export.csv
│   ├── wos_export.csv
│   └── sciencedirect_export.csv
│
└── scripts/
    ├── collect_arxiv.py
    ├── collect_openalex.py
    ├── collect_crossref.py
    ├── collect_github.py
    ├── clean_merge.py
    └── export_json.py
```

---

## 9. 如何更新資料

### 9.1 新增 Dataset

1. 在 `data/datasets.json` 新增 Dataset 資料。
2. 確認 `domain` 與 `subdomain` 名稱與 `subdomains.json` 一致。
3. 指定目前排名第一的 `best_model`。
4. 補上 `paper_url` 與 `github_url`。
5. 在 `data/papers.json` 新增該 Dataset 的 Ranking 資料。

### 9.2 新增 Paper / Model

1. 在 `data/papers.json` 新增一筆資料。
2. 確認 `domain`、`subdomain`、`dataset` 正確。
3. 填入 `rank`、`model_name`、`metrics`、`dataset_common_metrics`。
4. 補上 `paper_title`、`year_month`、`venue`、`status`。
5. 補上 `paper_url`、`github_url`、`source`。
6. 若資料尚未完全確認，`verified` 填 `Partial`。
7. 若 evaluation protocol 與其他方法不同，寫入 `notes`。

### 9.3 更新 Best Model

1. 確認該 Dataset 在 `papers.json` 中的 Rank 1。
2. 將 `datasets.json` 中該 Dataset 的 `best_model` 更新為 Rank 1 的 `model_name`。
3. 確保 Dataset card 顯示的 Best Model 與 See All 頁面中的 Rank 1 一致。

---

## 10. 如何部署到 GitHub Pages

### 10.1 建立 GitHub Repository

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

### 10.2 開啟 GitHub Pages

1. 進入 GitHub repository。
2. 點選 **Settings**。
3. 點選 **Pages**。
4. Source 選擇 `Deploy from a branch`。
5. Branch 選擇 `main`。
6. Folder 選擇 `/root`。
7. 儲存設定。

部署完成後，網站網址通常會是：

```text
https://<your-username>.github.io/<repo-name>/
```

---

## 11. 資料品質檢查清單

新增或更新資料前，請確認：

- [ ] 論文是真實存在的。
- [ ] Dataset 名稱正確。
- [ ] Ranking 只屬於單一 Dataset。
- [ ] Rank 依據明確。
- [ ] Metrics 有可追溯來源。
- [ ] Paper Link 可開啟。
- [ ] GitHub Link 為官方 repo 或已標註 unofficial。
- [ ] Year/Month 正確。
- [ ] Venue 正確。
- [ ] Status 正確。
- [ ] Source 已填寫。
- [ ] 不同 evaluation protocol 已在 notes 說明。
- [ ] Dataset card 的 Best Model 與該 Dataset Ranking 的 Rank 1 一致。

---

## 12. 開發優先順序

建議先完成：

1. 網站首頁與四大領域按鈕。
2. 三個已開放領域的子領域頁面。
3. Dataset card UI。
4. See All 單一 Dataset Ranking 頁面。
5. `papers.json`、`datasets.json`、`subdomains.json` 資料格式。
6. 每個子領域先放 3–5 個代表性 Dataset 示範資料。
7. 再逐步補齊 2021–2027 真實論文與排名資料。

---

## 13. 核心原則

本專案的核心原則：

```text
Dataset card 是入口。
See All 是單一 Dataset 的完整排名。
不同 Dataset 不混排。
不同 evaluation protocol 不強行比較。
所有資料都要可追溯。
```
