
---

## 1. 總搜尋 Prompt：系統性搜尋 IVAD 文獻

```text
I am conducting a systematic literature review on Industrial Visual Anomaly Detection (IVAD / IAD / VAD).

Please search across GitHub, arXiv, CVF Open Access, IEEE Xplore, Web of Science, Springer, ScienceDirect, OpenReview, Google Scholar, and Papers with Code for recent and influential papers.

Target datasets:
MVTec AD, VisA, MPDD, BTAD, MVTec AD 2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, MANTA, MAD-Sim, Eyecandies, DAGM, KolektorSDD, KSDD2.

Search terms must include:
"industrial anomaly detection", "visual anomaly detection", "industrial visual anomaly detection", "anomaly detection and localization", "surface defect detection", "visual defect detection", "unsupervised anomaly detection", "zero-shot anomaly detection", "few-shot anomaly detection", "anomaly segmentation", "defect localization", "industrial defect detection".

For each paper, extract:
1. Paper title
2. Year
3. Venue or preprint source
4. GitHub link, if available
5. Datasets used
6. Metrics reported: image AUROC, pixel AUROC, pixel AP, PRO, F1, mAP, FPS, throughput, inference time
7. Reported ranking or SOTA claim on each dataset
8. Method-based category:
   Reconstruction-based, Normalizing Flow-based, Representation-based, Data augmentation / anomaly synthesis-based, Student–Teacher-based, Knowledge Distillation-based, Memory Bank-based, Foundation Model-based, VLM / CLIP-based, Diffusion-based, 3D / RGB-D-based, Logical anomaly-based, Multi-view-based, Hybrid-based
9. Whether the result is from an official leaderboard, paper-reported table, GitHub reproduction, or third-party benchmark
10. Main contribution and limitation

Return the result as a table sorted by year and dataset.

Important:
Do not omit papers only because the title does not contain "industrial anomaly detection"; include papers that evaluate on MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, or MANTA.
Also identify which datasets have official leaderboards and which only have paper-reported comparisons.
Flag results that are not directly comparable because of different metrics, splits, input resolutions, backbones, pretraining settings, or hardware.
```

---

## 2. 找不同資料集排名 / SOTA 的 Prompt

```text
For each dataset in MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, and MANTA, find the most recent leaderboard or benchmark comparison table.

Separate the results into:
1. Official leaderboard results
2. Papers with Code leaderboard results
3. Paper-reported comparison tables
4. GitHub reproduction results

For each method, report:
method name, paper year, venue, dataset, image AUROC, pixel AUROC, pixel AP, PRO, FPS or inference speed, backbone, input resolution, whether code is available, and source link.

Important:
Do not merge rankings across datasets unless they use the same train/test split, same metric definition, and same evaluation protocol.
Flag results that are not directly comparable.
If FPS or throughput is reported, also record the GPU, batch size, and model size.
```

---

## 3. 找 2024–2026 最新 IVAD 方法的 Prompt

```text
Find industrial visual anomaly detection papers published or released from 2024 to 2026.

Focus on papers that evaluate on at least one of:
MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, MANTA.

Classify each paper into:
Reconstruction-based,
Normalizing Flow-based,
Representation-based,
Anomaly Synthesis-based,
Student–Teacher / Knowledge Distillation-based,
Memory Bank-based,
Foundation Model-based,
Vision-Language Model / CLIP-based,
Diffusion-based,
3D / RGB-D-based,
Logical anomaly-based,
Multi-view-based,
Hybrid / Unified-based.

For each paper, summarize:
1. What problem it solves
2. Which dataset challenge motivated it
3. What datasets it evaluates on
4. Whether it claims SOTA
5. Whether code is available
6. Its limitation
7. How it differs from EfficientAD, PatchCore, RD4AD, DRAEM, SimpleNet, GLASS, UniAD, UniNet, Dinomaly, Dinomaly2, AnomalyCLIP, WinCLIP, PromptAD
8. Whether the method is useful for improving EfficientAD
```

---

## 4. 找 Survey / Review / Taxonomy 的 Prompt

```text
Find recent survey papers from 2023 to 2026 about industrial visual anomaly detection, visual anomaly detection, industrial defect detection, anomaly segmentation, and anomaly localization.

Extract:
1. Paper title
2. Year
3. Venue
4. Taxonomy of methods
5. Taxonomy of datasets
6. Datasets discussed
7. Representative methods in each category
8. Open challenges mentioned
9. Future research directions
10. Whether the survey discusses foundation models, CLIP, DINOv2, SAM, diffusion models, logical anomaly detection, MVTec AD2, MVTec 3D-AD, or MVTec LOCO AD

Return a table and then summarize the most useful taxonomy for my literature review.
```

---

## 5. 找 GitHub 程式碼的 Prompt

```text
Search GitHub repositories related to industrial visual anomaly detection.

Target keywords:
MVTec AD anomaly detection,
VisA anomaly detection,
MPDD anomaly detection,
BTAD anomaly detection,
MVTec LOCO anomaly detection,
MVTec 3D anomaly detection,
DINOv2 industrial anomaly detection,
CLIP anomaly detection MVTec,
EfficientAD reproduction,
PatchCore reproduction,
Dinomaly2 GitHub,
GLASS anomaly detection,
UniNet anomaly detection.

For each GitHub repository, extract:
1. Repository name
2. URL
3. Related paper
4. Year
5. Supported datasets
6. Implemented methods
7. Reported metrics
8. Last update date
9. Stars
10. Whether training and inference code are available
11. Whether pretrained weights are provided
12. Whether evaluation scripts are available
13. Whether the implementation appears official or unofficial
```

---

## 6. Google / Google Scholar 搜尋式

```text
("industrial anomaly detection" OR "visual anomaly detection" OR "anomaly detection and localization" OR "surface defect detection" OR "defect localization")
AND
("MVTec AD" OR "VisA" OR "MPDD" OR "BTAD" OR "MVTec AD 2" OR "MVTec 3D" OR "MVTec LOCO" OR "Real-IAD")
```

```text
"industrial visual anomaly detection" "MVTec AD" "VisA" "pixel AUROC" "PRO"
```

```text
"visual anomaly detection" "foundation model" "DINOv2" "MVTec"
```

```text
"CLIP" "industrial anomaly detection" "MVTec AD" "VisA"
```

```text
"zero-shot anomaly detection" "MVTec AD" "VisA"
```

```text
"logical anomaly detection" "MVTec LOCO"
```

```text
"3D anomaly detection" "MVTec 3D-AD"
```

---

## 7. arXiv 搜尋式

```text
cat:cs.CV AND ("industrial anomaly detection" OR "visual anomaly detection" OR "anomaly detection and localization") AND ("MVTec" OR "VisA" OR "Real-IAD" OR "MVTec LOCO" OR "MVTec 3D")
```

```text
all:"MVTec AD" AND all:"VisA" AND all:"anomaly detection"
```

```text
all:"industrial anomaly detection" AND all:"DINOv2"
```

```text
all:"visual anomaly detection" AND all:"foundation model"
```

```text
all:"MVTec LOCO" AND all:"logical anomaly"
```

```text
all:"MVTec 3D" AND all:"anomaly detection"
```

```text
all:"Real-IAD" AND all:"visual anomaly detection"
```

```text
all:"industrial anomaly detection" AND submittedDate:[202401010000 TO 202612312359]
```

---

## 8. CVF Open Access 搜尋式

```text
site:openaccess.thecvf.com "MVTec AD" "anomaly detection"
```

```text
site:openaccess.thecvf.com "VisA" "anomaly detection"
```

```text
site:openaccess.thecvf.com "industrial anomaly detection" "MVTec"
```

```text
site:openaccess.thecvf.com "visual anomaly detection" "pixel AUROC"
```

```text
site:openaccess.thecvf.com "MVTec LOCO"
```

```text
site:openaccess.thecvf.com "MVTec 3D" "anomaly"
```

```text
site:openaccess.thecvf.com "foundation model" "anomaly detection" "MVTec"
```

---

## 9. IEEE Xplore 搜尋式

```text
("industrial anomaly detection" OR "visual anomaly detection" OR "surface defect detection" OR "industrial defect detection")
AND
("MVTec AD" OR "VisA" OR "MPDD" OR "BTAD" OR "MVTec LOCO" OR "MVTec 3D")
```

```text
(("industrial anomaly detection" OR "visual anomaly detection")
AND ("MVTec AD" OR "VisA" OR "Real-IAD"))
AND
(Publication Year: 2024 OR 2025 OR 2026)
```

```text
("industrial anomaly detection" OR "visual anomaly detection")
AND
("foundation model" OR "DINOv2" OR "CLIP" OR "vision-language model" OR "prompt learning")
AND
("MVTec" OR "VisA" OR "Real-IAD")
```

---

## 10. Web of Science 搜尋式

```text
TS=(("industrial anomaly detection" OR "visual anomaly detection" OR "surface defect detection" OR "defect localization" OR "anomaly segmentation")
AND
("MVTec AD" OR "VisA" OR "MPDD" OR "BTAD" OR "MVTec LOCO" OR "MVTec 3D" OR "Real-IAD"))
```

```text
TS=(("industrial anomaly detection" OR "visual anomaly detection")
AND
("MVTec AD" OR "VisA" OR "Real-IAD" OR "MVTec AD 2"))
AND PY=(2024 OR 2025 OR 2026)
```

```text
TS=(("industrial anomaly detection" OR "visual anomaly detection")
AND
("survey" OR "review" OR "benchmark" OR "taxonomy"))
```

---

## 11. 找 Foundation Model-based 方法的 Prompt

```text
Find all recent foundation model-based industrial visual anomaly detection papers.

Search terms:
DINOv2-based anomaly detection,
CLIP-based anomaly detection,
SAM-based anomaly detection,
vision-language anomaly detection,
prompt learning anomaly detection,
zero-shot industrial anomaly detection,
foundation model industrial anomaly detection.

Datasets:
MVTec AD, VisA, Real-IAD, MVTec AD2, MVTec LOCO, MANTA.

Return a table with:
paper, year, venue, model backbone, foundation model used, dataset, metrics, code link, method category, and limitation.

Also explain whether each method is:
DINOv2-based,
CLIP-based,
SAM-assisted,
VLM-based,
Prompt Learning-based,
Hybrid Foundation Model-based.
```

---

## 12. 找 Logical Anomaly-based 方法的 Prompt

```text
Find industrial anomaly detection papers focused on logical anomaly detection, structural anomaly detection, component-aware anomaly detection, relation-based anomaly detection, graph-based anomaly detection, and counting-based anomaly detection.

Datasets:
MVTec LOCO AD, VisA, MVTec AD2, Real-IAD.

Return:
1. Paper title
2. Year
3. Dataset
4. Method category
5. Whether it uses segmentation
6. Whether it uses graph reasoning
7. Whether it uses component modeling
8. Whether it uses global context
9. Whether it uses foundation models
10. Main contribution
11. Main limitation
```

---

## 13. 找 3D / RGB-D-based 方法的 Prompt

```text
Find 3D industrial anomaly detection papers using MVTec 3D-AD, RGB-D anomaly detection, point cloud anomaly detection, depth-map anomaly detection, or multi-modal RGB-depth fusion.

Return:
paper, year, venue, dataset, modality, method category, metrics, code link, and whether it uses RGB, depth, point cloud, or fusion.

Classify each method into:
3D-based,
RGB-D-based,
Point Cloud-based,
Depth Map-based,
Geometry Reconstruction-based,
Multi-modal Fusion-based,
Cross-modal Distillation-based,
Hybrid-based.
```

---

## 14. 找 Data Augmentation / Anomaly Synthesis-based 方法的 Prompt

```text
Find industrial visual anomaly detection papers that use data augmentation, synthetic anomaly generation, anomaly synthesis, CutPaste, DRAEM-like synthetic defects, GLASS-style anomaly synthesis, diffusion-generated anomalies, or procedural defect generation.

Datasets:
MVTec AD, VisA, MPDD, BTAD, Real-IAD, MVTec AD2.

For each paper, extract:
1. Paper title
2. Year
3. Dataset
4. Type of anomaly synthesis
5. Whether synthesis is image-level, patch-level, feature-level, or latent-level
6. Whether real anomaly samples are used
7. Metrics
8. Code link
9. Limitation
10. Whether it improves image-level AUROC, pixel-level AUROC, AP, or PRO
```

---

## 15. 找 Student–Teacher / Knowledge Distillation-based 方法的 Prompt

```text
Find industrial visual anomaly detection papers based on student-teacher learning, knowledge distillation, reverse distillation, teacher-student feature discrepancy, or feature mimicry.

Datasets:
MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec LOCO AD.

For each paper, extract:
1. Paper title
2. Year
3. Teacher backbone
4. Student architecture
5. Whether it uses autoencoder, projection head, bottleneck, memory bank, or anomaly synthesis
6. Datasets used
7. Image AUROC
8. Pixel AUROC
9. Pixel AP
10. PRO
11. FPS or inference time
12. Code link
13. Main contribution
14. Limitation

Compare each method with EfficientAD.
```

---

## 16. 找 Memory Bank / Nearest Neighbor-based 方法的 Prompt

```text
Find industrial visual anomaly detection papers based on memory bank, nearest neighbor search, patch-level feature retrieval, coreset selection, or feature embedding distance.

Datasets:
MVTec AD, VisA, MPDD, BTAD, Real-IAD.

For each paper, extract:
1. Paper title
2. Year
3. Feature extractor / backbone
4. Memory bank construction method
5. Distance metric
6. Coreset or feature selection strategy
7. Datasets
8. Metrics
9. Inference speed
10. Code link
11. Main contribution
12. Limitation
```

---

## 17. 文獻統計表建議欄位

| 欄位 | 說明 |
|---|---|
| Paper | 論文名稱 |
| Year | 年份 |
| Venue | CVPR / ICCV / ECCV / TIP / TII / arXiv 等 |
| Dataset | MVTec AD / VisA / MPDD / BTAD / MVTec AD2 / 3D / LOCO |
| Primary based | 主要方法類型 |
| Secondary based | 次要方法類型 |
| Dataset-driven based | 因資料集挑戰延伸出的方法類型 |
| Image AUROC | image-level 指標 |
| Pixel AUROC | pixel-level 指標 |
| Pixel AP | 定位精度 |
| PRO | anomaly region overlap |
| FPS / throughput | 推論速度 |
| GPU | RTX4090 / A100 / V100 等 |
| Code | GitHub 有無 |
| Ranking source | 官方 / Papers with Code / paper table / GitHub reproduction |
| Limitation | 方法限制 |
| Useful for EfficientAD? | 是否能補 EfficientAD 不足 |

---

## 18. 分階段搜尋流程

### Step 1：找 Survey / Taxonomy

```text
Find recent survey papers from 2023 to 2026 about industrial visual anomaly detection, visual anomaly detection, industrial defect detection, and anomaly localization. Extract their taxonomy of methods and datasets.
```

### Step 2：找 Benchmark / Leaderboard

```text
Find benchmark or leaderboard results for MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, and MANTA. Separate official leaderboards from paper-reported comparisons.
```

### Step 3：找最新 SOTA

```text
Find 2024–2026 papers claiming state-of-the-art or strong performance on industrial visual anomaly detection datasets. Include arXiv, CVPR, ICCV, ECCV, IEEE, Web of Science, GitHub, and OpenReview.
```

### Step 4：分類 Based

```text
Classify the collected papers by method-based category and dataset-driven category. Explain why each paper belongs to that category.
```

---

## 19. 最重要提醒

不要只搜尋：

```text
Find industrial visual anomaly detection papers.
```

更好的搜尋方式是：

```text
Search industrial visual anomaly detection papers using dataset names + method keywords + benchmark metrics + source-specific filters. Include MVTec AD, VisA, MPDD, BTAD, MVTec AD2, MVTec 3D-AD, MVTec LOCO AD, Real-IAD, and MANTA. Extract paper title, year, dataset, metrics, ranking source, GitHub link, FPS, and classify each method into reconstruction, NF, representation, anomaly synthesis, student–teacher, memory bank, foundation model, VLM, logical, 3D, multi-view, or hybrid-based categories.
```

---

## 20. 方法 based 分類參考

建議統計時至少保留以下 based：

1. Reconstruction-based
2. Normalizing Flow-based
3. Representation-based
4. Data Augmentation / Anomaly Synthesis-based
5. Student–Teacher-based
6. Knowledge Distillation-based
7. Memory Bank-based
8. Nearest Neighbor-based
9. One-class Classification-based
10. Contrastive Learning-based
11. Self-supervised Learning-based
12. Foundation Model-based
13. Vision-Language Model-based
14. Prompt Learning-based
15. Diffusion-based
16. Logical Anomaly Detection-based
17. Component Segmentation-based
18. Graph / Relation-based
19. Global-context-based
20. 3D / RGB-D-based
21. Point Cloud-based
22. Multi-modal Fusion-based
23. Multi-view-based
24. Domain Robustness-based
25. Noisy Normal-based
26. Tiny Defect-sensitive-based
27. Efficiency / Real-time-based
28. Hybrid / Unified-based
