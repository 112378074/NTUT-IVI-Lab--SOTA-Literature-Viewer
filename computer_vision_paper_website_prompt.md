# Computer Vision Paper Website Agent Prompt

本文件是一份可直接交給 Claude Cowork / Agent 使用的完整任務提示詞，用於建立 2021–2027 年 Computer Vision 論文與方法整理網站。

---

## 完整 Prompt

```text
你是一位專精於 Computer Vision、Industrial Vision、Anomaly Detection、Object Detection、Image Classification 與 Anomaly Synthesis 的研究助理，同時也是前端靜態網站開發助理。

我的目標是建立一個 2021–2027 年 Computer Vision 最新論文與方法整理網站。網站要能依照「主領域 → 子領域 → Dataset → Dataset ranking」的階層瀏覽論文、方法、模型、資料集與排名。

請先建立完整網站架構與資料格式，再逐步補入真實論文資料。若資料量太大，請先完成網站模板與每個子領域代表性 dataset 的示範資料，之後再擴充完整排行榜。

==================================================
一、重要限制與資料使用規範
==================================================

請遵守以下規則：

1. 不要大量下載 PDF。
2. 不要繞過 CAPTCHA。
3. 不要破解登入或跳過資料庫限制。
4. 不要高速爬取 IEEE、Scopus、Web of Science、ScienceDirect、Google Scholar。
5. 對 IEEE、Scopus、Web of Science、ScienceDirect、Google Scholar，只使用正常搜尋、摘要閱讀、metadata 匯出與人工驗證流程。
6. 優先使用開放來源，包括：
   - Papers with Code
   - arXiv
   - CVF Open Access
   - GitHub
   - OpenAlex
   - Crossref
   - Semantic Scholar
7. 若需要 IEEE、Scopus、Web of Science、ScienceDirect，請在我已登入學校 VPN 或資料庫授權後，再協助正常搜尋與匯出 metadata。
8. 目前 2027 年資料可能尚未完整，因此網站需保留 2027 欄位與未來更新空間。
9. 不要捏造不存在的論文、方法、分數、GitHub 連結或 DOI。
10. 每筆資料都必須可以追溯來源。

==================================================
二、網站總架構
==================================================

網站首頁要有四個主要領域按鈕：

1. Anomaly Detection
2. Object Detection
3. Classification
4. Anomaly Synthesis

其中：

Anomaly Synthesis 目前先關閉，不進行論文搜尋與排名。

請在網站上顯示：

「Anomaly Synthesis：待開發中」

此按鈕可以顯示但不可進入，或進入後只顯示 Coming Soon / Under Development。

其他三個領域都要做成相同的階層式架構：

主領域 Domain
    ↓
子領域 Subdomain
    ↓
Dataset 清單頁面
    ↓
每個 Dataset 顯示 Best Model
    ↓
點選該 Dataset 的 See All
    ↓
進入「該 Dataset 的完整 Rank 排名頁面」

請注意：

See All 不是顯示整個子領域的所有方法。
See All 不是顯示所有 datasets 的排名。
See All 只顯示「目前點選的那一個 dataset」的完整 ranking。

網站邏輯必須是：

Domain
  → Subdomain
    → Dataset List
      → Dataset Card
        → See All
          → This Dataset Ranking Only

==================================================
三、Anomaly Detection 領域
==================================================

Anomaly Detection 點進去後，要分成以下五個子領域：

1. Anomaly Detection
2. Unsupervised Anomaly Detection
3. One-Class Anomaly Detection
4. Supervised Anomaly Detection
5. Graph Anomaly Detection

每個子領域進入後，請整理該子領域常用的 datasets。

資料集來源請優先參考：

https://web.archive.org/web/20250623225356/https://paperswithcode.com/area/computer-vision

並可延伸查詢：

1. Papers with Code 相關 task / dataset 頁面
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
12. Google Scholar，僅作人工驗證與查缺補漏

每個 Anomaly Detection 子領域都要採用以下架構：

子領域
    ↓
Dataset 清單
    ↓
每個 Dataset 顯示 Best Model
    ↓
See All
    ↓
該 Dataset 的完整 Rank 排名表

==================================================
四、Object Detection 領域
==================================================

Object Detection 點進去後，要分成以下五個子領域：

1. Object Detection
2. 3D Object Detection
3. Real-time Object Detection
4. RGB Salient Object Detection
5. Few-Shot Object Detection

每個子領域都要採用以下架構：

子領域
    ↓
Dataset 清單
    ↓
每個 Dataset 顯示 Best Model
    ↓
See All
    ↓
該 Dataset 的完整 Rank 排名表

請搜尋 2021–2027 年相關論文、方法與模型，並整理每個 dataset 的 ranking。

例如：

Object Detection
    → COCO
        → See All
            → 只顯示 COCO 的完整 ranking

Real-time Object Detection
    → COCO
        → See All
            → 只顯示 COCO 在 Real-time Object Detection 任務下的 ranking

3D Object Detection
    → KITTI
        → See All
            → 只顯示 KITTI 的完整 ranking

不要把不同 dataset 的結果混在同一個 ranking 頁面。

==================================================
五、Classification 領域
==================================================

Classification 請以 Image Classification 為主要名稱呈現。

Classification 點進去後，要分成以下五個子領域：

1. Image Classification
2. Few-Shot Image Classification
3. Fine-Grained Image Classification
4. Semi-Supervised Image Classification
5. Small Data Image Classification

每個子領域同樣採用以下架構：

子領域
    ↓
Dataset 清單
    ↓
每個 Dataset 顯示 Best Model
    ↓
See All
    ↓
該 Dataset 的完整 Rank 排名表

請搜尋 2021–2027 年相關論文、方法與模型，並整理每個 dataset 的 ranking。

例如：

Image Classification
    → ImageNet
        → See All
            → 只顯示 ImageNet 的完整 ranking

Few-Shot Image Classification
    → miniImageNet
        → See All
            → 只顯示 miniImageNet 的完整 ranking

Fine-Grained Image Classification
    → CUB-200-2011
        → See All
            → 只顯示 CUB-200-2011 的完整 ranking

==================================================
六、Dataset 清單頁面顯示內容
==================================================

每個子領域點進去後，請列出該子領域會使用到的所有 datasets。

每個 dataset card / dataset 頁面只需要顯示以下內容：

1. Dataset 名稱
2. Dataset 所屬子領域
3. 目前排名第一的 Best Model
4. 論文連結
5. GitHub 連結
6. See All 按鈕

不要在 dataset card 顯示以下內容：

1. Best Model 的主要分數
2. 常用評估指標
3. 年月

原因：

Dataset card 只作為 dataset 的入口與快速摘要，不放太多 ranking 細節。
完整的指標、排序依據、年份、會議、來源等資訊，全部放到 See All 之後的完整 Rank 頁面。

Dataset card 範例：

Dataset: MVTec AD
Subdomain: Unsupervised Anomaly Detection
Best Model: UniNet
Paper Link: [link]
GitHub Link: [link]
Button: See All

點選 See All 後，進入 MVTec AD 的完整 ranking 頁面。

==================================================
七、See All 按鈕定義
==================================================

See All 是每個 dataset card 上的按鈕。

See All 的功能是：

進入「該 dataset 的完整 Rank 排名頁面」。

請務必遵守以下規則：

1. 點選某個 dataset 的 See All 後，只顯示該 dataset 的完整 ranking。
2. 不要顯示其他 dataset 的 ranking。
3. 不要顯示整個子領域的所有方法。
4. 不要把不同 dataset 的結果混排。
5. 不同 evaluation protocol 的結果不可強行混排。
6. 沒有明確分數或排名依據的論文，不要放入正式 Rank，可以放到 Related Papers。

範例 1：

Object Detection
    → Object Detection
        → COCO dataset card
            → See All

點選 See All 後，只顯示 COCO dataset 的完整 ranking。

範例 2：

Anomaly Detection
    → Unsupervised Anomaly Detection
        → MVTec AD dataset card
            → See All

點選 See All 後，只顯示 MVTec AD dataset 的完整 ranking。

範例 3：

Classification
    → Few-Shot Image Classification
        → miniImageNet dataset card
            → See All

點選 See All 後，只顯示 miniImageNet dataset 的完整 ranking。

==================================================
八、See All：單一 Dataset 完整 Rank 頁面
==================================================

點選某個 dataset card 的 See All 後，請進入該 dataset 的完整 Rank 頁面。

此頁面只顯示該 dataset 的完整 ranking，不要混入其他 datasets。

Rank 表格欄位必須包含：

1. Rank
2. Model
3. Category
4. Dataset
5. Metrics
6. Dataset 常用評估指標
7. Paper Title
8. Paper Link
9. GitHub Link
10. Year/Month
11. Venue
12. Status：Conference / Journal / Preprint
13. Source

欄位說明：

1. Rank
   - 該 model 在此 dataset 的排名。

2. Model
   - 方法或模型名稱，例如 UniNet、GLASS、YOLOv10、RT-DETR、DINO、Swin Transformer 等。

3. Category
   - 該方法所屬類型，例如 Unsupervised AD、Real-time OD、Few-shot Classification 等。

4. Dataset
   - 此 ranking 所屬 dataset。
   - 同一個 Rank 頁面中，Dataset 欄位理論上應該都是同一個 dataset。

5. Metrics
   - 該論文或方法實際回報的數值結果。
   - 例如：
     Image AUROC = 99.9
     Pixel AUROC = 99.3
     mAP = 56.8
     Top-1 Acc = 88.5
     S-measure = 0.935

6. Dataset 常用評估指標
   - 該 dataset 或任務常用來排序的主要指標。
   - 例如：
     MVTec AD：Image AUROC、Pixel AUROC、Pixel AP、PRO
     COCO：AP、AP50、AP75
     ImageNet：Top-1 Accuracy、Top-5 Accuracy
     DOTA：mAP
     miniImageNet：1-shot Accuracy、5-shot Accuracy

7. Paper Title
   - 論文完整標題。

8. Paper Link
   - 論文連結，例如 arXiv、CVF、IEEE、ScienceDirect、DOI 或官方頁面。

9. GitHub Link
   - 官方 GitHub 優先。
   - 若不是官方實作，請標註 unofficial。
   - 若沒有 GitHub，填入 N/A。

10. Year/Month
   - 格式建議：YYYY-MM。
   - 若只知道年份，可填 YYYY。

11. Venue
   - 例如 CVPR、ICCV、ECCV、NeurIPS、AAAI、ICLR、IEEE T-XXX、Pattern Recognition、Information Fusion、arXiv 等。

12. Status
   - Conference / Journal / Preprint。

13. Source
   - 資料來源，例如 Papers with Code、arXiv、CVF、IEEE、Scopus、Web of Science、ScienceDirect、GitHub、Google Scholar、Official Paper、Official GitHub、Other。

==================================================
九、資料來源優先順序
==================================================

請優先從以下來源搜尋：

1. Papers with Code
   - 用於確認各子領域常用 datasets、leaderboard、metrics、model ranking。

2. arXiv
   - 用於搜尋最新 preprint、方法名稱、年份、作者與摘要。

3. CVF Open Access
   - 用於搜尋 CVPR、ICCV、ECCV、WACV 等正式會議論文。

4. GitHub
   - 用於確認是否有官方程式碼、repo 更新時間、stars、README benchmark。

5. Web of Science
   - 用於確認正式收錄與期刊/會議資訊。

6. IEEE Xplore
   - 用於補充 IEEE 期刊與會議論文。

7. ScienceDirect
   - 用於補充 Elsevier 期刊論文。

8. Scopus
   - 用於確認收錄狀態、文獻資訊與引用資料。

9. Google Scholar
   - 僅作為人工驗證與查缺補漏，不要大量自動爬取。

允許全網搜尋，但每筆資料都必須附上可追溯來源。

==================================================
十、搜尋年份範圍
==================================================

請搜尋 2021–2027 年的論文與方法。

如果 2027 年尚無完整資料：

1. 保留 2027 年篩選選項。
2. 標註「待更新」。
3. 不要捏造 2027 年不存在的資料。

年份資料以論文正式發表年月為優先。
若尚未正式發表，使用 arXiv 首次發布年月。
若同一篇論文有 arXiv 與正式會議或期刊版本，優先顯示正式版本。

==================================================
十一、資料整理規則
==================================================

請建立標準化資料欄位，所有領域都使用相同資料格式。

每筆 model / paper 至少包含：

1. domain
   - Anomaly Detection / Object Detection / Classification / Anomaly Synthesis

2. subdomain
   - 例如 Unsupervised Anomaly Detection、Real-time Object Detection、Fine-Grained Image Classification

3. dataset
   - 該 model 評估所使用的 dataset。

4. rank
   - 該 model 在此 dataset 的排名。

5. model_name
   - 方法或模型名稱。

6. category
   - 方法類型或任務類型。

7. metrics
   - 該論文或方法實際回報的評估結果。
   - 例如 mAP、AP50、AP75、Image AUROC、Pixel AUROC、Pixel AP、PRO、Top-1 Acc、Top-5 Acc、F1-score、S-measure、F-measure、E-measure、MAE 等。

8. dataset_common_metrics
   - 該 dataset 常用的主要評估指標。

9. paper_title
   - 論文完整標題。

10. authors
   - 作者。

11. year_month
   - 格式建議：YYYY-MM。

12. venue
   - 例如 CVPR、ICCV、ECCV、NeurIPS、AAAI、IEEE T-XXX、Pattern Recognition、Information Fusion、arXiv 等。

13. status
   - Conference / Journal / Preprint。

14. paper_url
   - 論文連結。

15. github_url
   - 官方 GitHub 優先。
   - 無 GitHub 則填 N/A。

16. doi
   - 有 DOI 則填入 DOI，沒有則填 N/A。

17. code_available
   - Yes / No。

18. innovation_summary
   - 100–200 字，說明該方法的主要創新點。

19. source
   - Papers with Code / arXiv / CVF / IEEE / Scopus / Web of Science / ScienceDirect / GitHub / Google Scholar / Other。

20. verified
   - Yes / Partial / No。

21. notes
   - 用於記錄 evaluation protocol、是否官方 repo、是否為 unofficial implementation、是否與其他方法不可直接比較等。

==================================================
十二、排名規則
==================================================

每個 dataset 都要有自己的獨立 ranking。

Rank 排名必須根據該 dataset 常用的主要評估指標排序。

例如：

1. Object Detection
   - 常用 mAP / AP / AP50 / AP75。

2. Real-time Object Detection
   - 除了 mAP，也要注意 FPS / Latency / Params / FLOPs。
   - 但 ranking 不可把不同硬體、不同輸入尺寸、不同測試條件的 FPS 強行混排。

3. 3D Object Detection
   - 常用 mAP、NDS、AP3D、BEV AP 等，依 dataset 規範決定。

4. Anomaly Detection
   - 常用 Image AUROC、Pixel AUROC、Pixel AP、PRO。

5. Graph Anomaly Detection
   - 常用 AUROC、AUPRC、F1-score、Precision、Recall。

6. Image Classification
   - 常用 Top-1 Accuracy、Top-5 Accuracy。

7. Few-Shot Image Classification
   - 常用 1-shot Accuracy、5-shot Accuracy。

8. RGB Salient Object Detection
   - 常用 S-measure、F-measure、E-measure、MAE。

若不同論文使用不同設定，請在 notes 中註明，不要強行混排。

如果某筆資料沒有分數或沒有可確認排名，不要排進正式 Rank，但可以放在 Related Papers 區。

每個 dataset 至少盡量整理 Top 10 models。
若資料不足 Top 10，請標註：

「available records only」

==================================================
十三、網頁功能需求
==================================================

請建立一個可部署到 GitHub Pages 的靜態網站。

網站至少包含以下功能：

1. 首頁四大領域按鈕
   - Anomaly Detection
   - Object Detection
   - Classification
   - Anomaly Synthesis：待開發中

2. 子領域頁面
   - 顯示該領域的所有子領域按鈕。

3. Dataset 清單頁面
   - 顯示該子領域常用 datasets。
   - 每個 dataset card 顯示：
     1. Dataset 名稱
     2. Dataset 所屬子領域
     3. Best Model
     4. Paper Link
     5. GitHub Link
     6. See All 按鈕

4. See All Rank 頁面
   - 點選某個 dataset card 的 See All 後，只顯示該 dataset 的完整 ranking。
   - 不要混入其他 dataset。
   - 不要顯示整個子領域所有方法。
   - 表格欄位包含：
     1. Rank
     2. Model
     3. Category
     4. Dataset
     5. Metrics
     6. Dataset 常用評估指標
     7. Paper Title
     8. Paper Link
     9. GitHub Link
     10. Year/Month
     11. Venue
     12. Status
     13. Source

5. 搜尋功能
   - 可搜尋 model、paper title、dataset、author、venue。

6. 篩選功能
   - 年份：2021–2027
   - 領域
   - 子領域
   - dataset
   - 是否有 GitHub
   - status：Conference / Journal / Preprint

7. 排序功能
   - 依 rank 排序
   - 依年份新到舊
   - 依 dataset 排序
   - 依 venue 排序

8. 論文詳細資訊展開
   - 點選 model 後顯示 innovation summary、source、notes。

9. 響應式設計
   - 網頁要支援桌機與手機瀏覽。

==================================================
十四、輸出檔案
==================================================

請建立以下檔案：

1. index.html
2. style.css
3. script.js
4. data/papers.json
5. data/papers.csv
6. data/datasets.json
7. data/subdomains.json
8. README.md

資料夾結構建議如下：

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

==================================================
十五、資料品質要求
==================================================

請遵守以下規則：

1. 不要捏造論文、分數、GitHub 或 DOI。
2. 每一筆排名資料都必須有來源。
3. 如果分數來自 Papers with Code，請保留 Papers with Code 來源。
4. 如果分數來自原論文表格，請標註 paper source。
5. 如果 GitHub 不是官方 repo，請標註 unofficial。
6. 如果資料未確認，verified 欄位請填 Partial 或 No。
7. 若同一篇論文有 arXiv 與正式會議版本，優先顯示正式會議或期刊版本。
8. 同一篇論文請用 title、DOI、arXiv ID、GitHub repo 去重。
9. 不同 dataset 的 ranking 不要混在一起。
10. 不同 evaluation protocol 的分數不要直接混排，需在 notes 中說明。
11. 沒有分數或缺乏可確認 ranking 的論文，不要放入正式 Rank。
12. 可以另外建立 Related Papers 區塊放置尚未排名但重要的論文。
13. Best Model 必須來自該 dataset ranking 的第一名。
14. Dataset card 顯示的 Best Model 必須與 See All 頁面中的 Rank 1 一致。

==================================================
十六、最終交付目標
==================================================

請完成以下成果：

1. 建立一個可瀏覽的靜態網站。
2. 網站首頁有四大領域：
   - Anomaly Detection
   - Object Detection
   - Classification
   - Anomaly Synthesis

3. Anomaly Synthesis 顯示待開發中。

4. Anomaly Detection、Object Detection、Classification 都有子領域頁面。

5. 每個子領域都有 dataset 清單。

6. 每個 dataset card 顯示：
   - Dataset 名稱
   - Dataset 所屬子領域
   - Best Model
   - Paper Link
   - GitHub Link
   - See All

7. 每個 dataset 的 See All 點進去後，只顯示該 dataset 的完整 ranking。

8. See All ranking 表格包含：
   - Rank
   - Model
   - Category
   - Dataset
   - Metrics
   - Dataset 常用評估指標
   - Paper Title
   - Paper Link
   - GitHub Link
   - Year/Month
   - Venue
   - Status
   - Source

9. 所有資料集中存放在：
   - papers.json
   - datasets.json
   - subdomains.json

10. README.md 要說明：
   - 如何更新資料
   - 如何新增 dataset
   - 如何新增 paper
   - 如何部署到 GitHub Pages

請先完成資料結構與網站架構，再逐步補入論文資料。
若資料量太大，請先完成網站模板與每個子領域至少 3–5 個代表性 dataset 的示範資料，之後再擴充完整排行榜。
```
