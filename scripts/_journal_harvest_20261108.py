#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Weekly IVAD Journal Harvest — 2026-11-08
Coverage window: 2026-11-02 to 2026-11-08

Target journals searched:
  IEEE TPAMI, IEEE TII, IEEE TNNLS, IEEE TIP, IEEE TCSVT,
  Pattern Recognition, IJCV, Computers in Industry,
  Information Fusion, Knowledge-Based Systems,
  Information Sciences, Neural Networks

Network access this run:
  - arXiv.org: BLOCKED by environment proxy (ENOTFOUND)
  - IEEE Xplore: BLOCKED by environment proxy
  - ScienceDirect: BLOCKED by environment proxy
  - ResearchGate: BLOCKED by environment proxy
  - alphaxiv.org, emergentmind.com, huggingface.co: BLOCKED
  - Web search (Google snippets): ACCESSIBLE

Dupe-check results (against workbook):
  - IGAF (Information Fusion): ALREADY IN WORKBOOK (prior harvest)
  - INP-Former / INP-Former++: ALREADY IN WORKBOOK
  - KairosAD: ICIAP 2025 conference (not target journal), skip
  - arXiv 2606.07659: supervised YOLO OD on NEU-DET, not anomaly detection, skip
  - Bounded AUKS (2611.04754): evaluation metric paper, not a method — skip
  - INP-Former reproduction (2611.01230): reproduction study, not new method — skip

New entries to add this week: 3 papers

1. DART (arXiv 2611.08271, Nov 3–8, 2026)
   "DART: Domain-Agnostic Residual Transfer for Generalist Anomaly Detection"
   – Nov 2026 arXiv preprint; title confirmed from Google index of arxiv.org/html/2611.08271.
   – Generalist AD method evaluated on MVTec AD + VisA (+ BraTS as cross-domain).
   – arXiv HTML blocked; no GitHub repo accessible (404).
   – Relative improvement reported (+1.15 on MVTec AD, +0.68 on VisA over baseline)
     in a web-search AI snippet — NOT from an authoritative table; DO NOT use for ranking.
   – Staged ⏳ Pending with blank metrics.

2. SAM-IAD (Knowledge-Based Systems, vol. 338, art. 115515, Feb 2026)
   "SAM-IAD: Injecting specific knowledge into SAM for industrial anomaly detection"
   – Authors: Yichi Chen, Bin Chen, Weizhi Xian, Jian Wang, Xiaoping Gong, Jing Han,
     Xiaohua Tao (Harbin Institute of Technology)
   – Published: Knowledge-Based Systems, vol. 338, art. 115515, February 2026.
     DOI: 10.1016/j.knosys.2026.115515
   – Datasets: MVTec AD, VisA, DAGM. Claims SOTA in anomaly segmentation.
   – Method: anomaly knowledge injection into frozen SAM via cross-color-space residuals
     + learnable anomaly adapters in image encoder. No arXiv preprint found.
   – No GitHub found. ScienceDirect blocked → no accessible metrics table.
   – Missed by all prior harvests. Staged ⏳ Pending.

3. DiffAST (arXiv 2604.19240, Apr 21, 2026)
   "Industrial Surface Defect Detection via Diffusion Generation and Asymmetric Student-Teacher Network"
   – Authors: Shuo Feng, Runlin Zhou, Yuyang Li, Guangcan Liu (Southeast University)
   – No confirmed journal submission (pure arXiv preprint as of search date).
   – Datasets: MVTec AD, VisA, MPDD.
   – Method: DDPM trained on normal samples to synthesize defects with pixel masks;
     frozen teacher + student reconstructing normal features; cosine similarity + seg loss.
   – Abstract states MVTec AD I-AUROC 98.4, P-AUROC 98.3; third-party summary cites
     VisA I-AUROC 92.3 and MPDD I-AUROC 95.8. arXiv HTML blocked; no GitHub found.
     Abstract-level claim NOT sufficient for ranking → staged ⏳ Pending.
   – Missed by all prior harvests.

No SMTP credentials found (scripts/.env absent) → email skipped.
"""

from pathlib import Path
import sys

try:
    import openpyxl
except ImportError:
    # Try system python3.13 site-packages path
    sys.path.insert(0, '/usr/local/lib/python3.13/site-packages')
    import openpyxl

AD_XLSX = Path(__file__).resolve().parent.parent / 'AnomalyDetection_Papers_Summary_v10_20260425.xlsx'

# ---------------------------------------------------------------------------
# Row data — column order (15 cols):
# (資料集, 類別, 方法, 作者, 發表, 年月, 狀態,
#  I-AUROC, P-AUROC, P-AP, P-PRO, FPS, 備註(特色/based), 連結, GitHub)
# ---------------------------------------------------------------------------

# ============================================================
# 1. DART (arXiv 2611.08271, Nov 2026)
#    "DART: Domain-Agnostic Residual Transfer for Generalist Anomaly Detection"
# ============================================================

DART_OVERVIEW = (
    'MVTec AD, VisA (+ BraTS cross-domain; Generalist AD)',
    '表徵學習 (Representation-based) / 殘差遷移 (Residual Transfer-based) '
    '/ 通用型異常偵測 (Generalist AD)',
    'DART (Domain-Agnostic Residual Transfer for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.08271)',
    'arXiv 2611.08271 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08271 (Nov 2026); arXiv HTML blocked, no accessible GitHub. '
    'Title confirmed from Google index of arxiv.org/html/2611.08271. '
    'Generalist anomaly detection via domain-agnostic residual transfer. '
    'Evaluates on MVTec AD + VisA as primary industrial benchmarks; also evaluates on '
    'BraTS (medical) as cross-domain test. Relative improvement reported as +1.15 '
    'on MVTec AD and +0.68 on VisA over baseline (source: web-search snippet — NOT '
    'from authoritative table; absolute AUROC values inaccessible). '
    'Journal harvest 2026-11-08. [Representation / Residual Transfer / Generalist AD]',
    'https://arxiv.org/abs/2611.08271',
    None,
)

DART_MVTEC = (
    'MVTec AD',
    '表徵學習 (Representation-based) / 殘差遷移 (Residual Transfer-based) '
    '/ 通用型異常偵測 (Generalist AD)',
    'DART (Domain-Agnostic Residual Transfer for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.08271)',
    'arXiv 2611.08271 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08271 (Nov 2026); arXiv HTML blocked. '
    'Relative improvement on MVTec AD: +1.15 over baseline (web snippet, NOT authoritative table). '
    'Absolute I-AUROC inaccessible. [Representation / Residual Transfer / Generalist AD]',
    'https://arxiv.org/abs/2611.08271',
    None,
)

DART_VISA = (
    'VisA',
    '表徵學習 (Representation-based) / 殘差遷移 (Residual Transfer-based) '
    '/ 通用型異常偵測 (Generalist AD)',
    'DART (Domain-Agnostic Residual Transfer for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.08271)',
    'arXiv 2611.08271 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08271 (Nov 2026); arXiv HTML blocked. '
    'Relative improvement on VisA: +0.68 over baseline (web snippet, NOT authoritative table). '
    'Absolute I-AUROC inaccessible. [Representation / Residual Transfer / Generalist AD]',
    'https://arxiv.org/abs/2611.08271',
    None,
)

# ============================================================
# 2. SAM-IAD (Knowledge-Based Systems, vol. 338, art. 115515, Feb 2026)
#    "SAM-IAD: Injecting specific knowledge into SAM for industrial anomaly detection"
# ============================================================

SAMIAD_OVERVIEW = (
    'MVTec AD, VisA, DAGM',
    '基礎模型 (Foundation Model-based) / SAM-based '
    '/ 知識注入 (Knowledge Injection-based) / 異常適配器 (Anomaly Adapter-based)',
    'SAM-IAD (Injecting specific knowledge into SAM for industrial anomaly detection)',
    'Yichi Chen, Bin Chen, Weizhi Xian, Jian Wang, Xiaoping Gong, Jing Han, Xiaohua Tao '
    '(Harbin Institute of Technology)',
    'Knowledge-Based Systems, vol. 338, art. 115515 (2026-02)',
    '2026-02',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — Knowledge-Based Systems vol. 338, art. 115515, Feb 2026; '
    'DOI: 10.1016/j.knosys.2026.115515. No arXiv preprint found. No GitHub found. '
    'ScienceDirect blocked — no authoritative table accessible. '
    'Method: Injects anomaly knowledge into frozen SAM via cross-color-space residual mapping '
    'as mask prompts + learnable anomaly adapters in image encoder. Evaluates on MVTec AD, '
    'VisA, and DAGM; claims SOTA in anomaly segmentation. '
    'Target journal: Knowledge-Based Systems. Missed by all prior harvests. '
    'Journal harvest 2026-11-08. [Foundation Model / SAM / Knowledge Injection / Anomaly Adapter]',
    'https://www.sciencedirect.com/science/article/abs/pii/S0950705126002571',
    None,
)

SAMIAD_MVTEC = (
    'MVTec AD',
    '基礎模型 (Foundation Model-based) / SAM-based / 知識注入 (Knowledge Injection-based)',
    'SAM-IAD (Injecting specific knowledge into SAM for industrial anomaly detection)',
    'Yichi Chen, Bin Chen, Weizhi Xian, Jian Wang, Xiaoping Gong, Jing Han, Xiaohua Tao (HIT)',
    'Knowledge-Based Systems, vol. 338, art. 115515 (2026-02)',
    '2026-02',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — KBS vol. 338 art. 115515, Feb 2026; ScienceDirect blocked; '
    'no accessible metrics table (no arXiv, no GitHub). '
    'Claims SOTA anomaly segmentation on MVTec AD. '
    '[Foundation Model / SAM / Knowledge Injection]',
    'https://www.sciencedirect.com/science/article/abs/pii/S0950705126002571',
    None,
)

SAMIAD_VISA = (
    'VisA',
    '基礎模型 (Foundation Model-based) / SAM-based / 知識注入 (Knowledge Injection-based)',
    'SAM-IAD (Injecting specific knowledge into SAM for industrial anomaly detection)',
    'Yichi Chen, Bin Chen, Weizhi Xian, Jian Wang, Xiaoping Gong, Jing Han, Xiaohua Tao (HIT)',
    'Knowledge-Based Systems, vol. 338, art. 115515 (2026-02)',
    '2026-02',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — KBS vol. 338 art. 115515, Feb 2026; ScienceDirect blocked; '
    'no accessible metrics table (no arXiv, no GitHub). '
    'Claims SOTA anomaly segmentation on VisA. '
    '[Foundation Model / SAM / Knowledge Injection]',
    'https://www.sciencedirect.com/science/article/abs/pii/S0950705126002571',
    None,
)

# ============================================================
# 3. DiffAST (arXiv 2604.19240, Apr 21, 2026)
#    "Industrial Surface Defect Detection via Diffusion Generation
#     and Asymmetric Student-Teacher Network"
# ============================================================

DIFFAST_OVERVIEW = (
    'MVTec AD, VisA, MPDD',
    '擴散模型 (Diffusion-based) / 師生網路 (Student-Teacher-based) '
    '/ 異常合成 (Anomaly Synthesis-based) / 非對稱師生 (Asymmetric ST-based)',
    'DiffAST (Industrial Surface Defect Detection via Diffusion Generation '
    'and Asymmetric Student-Teacher Network)',
    'Shuo Feng, Runlin Zhou, Yuyang Li, Guangcan Liu (Southeast University, Nanjing, China)',
    'arXiv 2604.19240 (2026-04-21)',
    '2026-04',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2604.19240 (Apr 21, 2026); arXiv HTML blocked; no GitHub found. '
    'Method: DDPM trained on normal samples only to synthesize anomalous images with pixel-level '
    'masks; frozen teacher provides stable normal feature baseline; student reconstructs normal '
    'features from anomalous inputs; loss combines cosine similarity and segmentation supervision. '
    'Evaluates on MVTec AD, VisA, and MPDD. Abstract states MVTec AD I-AUROC 98.4 / P-AUROC 98.3. '
    'VisA I-AUROC 92.3 and MPDD I-AUROC 95.8 cited from third-party review (themoonlight.io) — '
    'NOT from official repo or arXiv HTML table; NOT ranked pending direct verification. '
    'No confirmed journal venue. Missed by prior harvests. '
    'Journal harvest 2026-11-08. [Diffusion / Student-Teacher / Anomaly Synthesis / Asymmetric ST]',
    'https://arxiv.org/abs/2604.19240',
    None,
)

DIFFAST_MVTEC = (
    'MVTec AD',
    '擴散模型 (Diffusion-based) / 師生網路 (Student-Teacher-based) '
    '/ 異常合成 (Anomaly Synthesis-based)',
    'DiffAST (Industrial Surface Defect Detection via Diffusion Generation '
    'and Asymmetric Student-Teacher Network)',
    'Shuo Feng, Runlin Zhou, Yuyang Li, Guangcan Liu (Southeast University)',
    'arXiv 2604.19240 (2026-04-21)',
    '2026-04',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2604.19240 (Apr 2026); arXiv HTML blocked; no GitHub found. '
    'Abstract states MVTec AD I-AUROC 98.4, P-AUROC 98.3. NOT ranked — requires arXiv HTML '
    'or official repo table verification before ranking. '
    '[Diffusion / Student-Teacher / Anomaly Synthesis]',
    'https://arxiv.org/abs/2604.19240',
    None,
)

DIFFAST_VISA = (
    'VisA',
    '擴散模型 (Diffusion-based) / 師生網路 (Student-Teacher-based) '
    '/ 異常合成 (Anomaly Synthesis-based)',
    'DiffAST (Industrial Surface Defect Detection via Diffusion Generation '
    'and Asymmetric Student-Teacher Network)',
    'Shuo Feng, Runlin Zhou, Yuyang Li, Guangcan Liu (Southeast University)',
    'arXiv 2604.19240 (2026-04-21)',
    '2026-04',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2604.19240 (Apr 2026); arXiv HTML blocked; no GitHub found. '
    'Third-party review cites VisA I-AUROC 92.3, P-AUROC 97.4 (NOT authoritative). '
    'NOT ranked pending verification from official table. '
    '[Diffusion / Student-Teacher / Anomaly Synthesis]',
    'https://arxiv.org/abs/2604.19240',
    None,
)

DIFFAST_MPDD = (
    'MPDD',
    '擴散模型 (Diffusion-based) / 師生網路 (Student-Teacher-based) '
    '/ 異常合成 (Anomaly Synthesis-based)',
    'DiffAST (Industrial Surface Defect Detection via Diffusion Generation '
    'and Asymmetric Student-Teacher Network)',
    'Shuo Feng, Runlin Zhou, Yuyang Li, Guangcan Liu (Southeast University)',
    'arXiv 2604.19240 (2026-04-21)',
    '2026-04',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2604.19240 (Apr 2026); arXiv HTML blocked; no GitHub found. '
    'Third-party review cites MPDD I-AUROC 95.8, P-AUROC 95.6 (NOT authoritative). '
    'NOT ranked pending verification from official table. '
    '[Diffusion / Student-Teacher / Anomaly Synthesis]',
    'https://arxiv.org/abs/2604.19240',
    None,
)


# ---------------------------------------------------------------------------
# Insert helpers
# ---------------------------------------------------------------------------

def sheet_has_link(ws, link):
    """Dupe check: look for identical link in column 14 (index 13)."""
    link_lower = link.lower()
    for row in ws.iter_rows(values_only=True):
        if len(row) > 13 and row[13] and str(row[13]).lower() == link_lower:
            return True
    return False


def safe_append(ws, row_tuple):
    link = row_tuple[13]
    if link and sheet_has_link(ws, link):
        # Distinguish rows that share the same arXiv/DOI URL by checking method name too
        method_name = str(row_tuple[2])[:40].lower()
        for row in ws.iter_rows(values_only=True):
            if (len(row) > 13 and row[13] and
                    str(row[13]).lower() == link.lower() and
                    len(row) > 2 and row[2] and
                    method_name in str(row[2]).lower()):
                print(f'  SKIP (dupe by URL+method): {str(row_tuple[2])[:55]} → {ws.title}')
                return False
    ws.append(row_tuple)
    print(f'  ADD: {str(row_tuple[2])[:55]} → {ws.title}')
    return True


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))

    # Identify sheets
    ws_all   = None
    ws_mvtec = wb['MVTec AD']
    ws_visa  = wb['VisA']
    ws_mpdd  = wb['MPDD']

    for name in wb.sheetnames:
        if 'All' in name or '總' in name or '総' in name:
            ws_all = wb[name]
            break
    if ws_all is None:
        ws_all = wb.active

    added = 0

    print('=== Adding DART (arXiv 2611.08271) ===')
    added += safe_append(ws_all,   DART_OVERVIEW)
    added += safe_append(ws_mvtec, DART_MVTEC)
    added += safe_append(ws_visa,  DART_VISA)

    print('=== Adding SAM-IAD (KBS vol. 338, 2026) ===')
    added += safe_append(ws_all,   SAMIAD_OVERVIEW)
    added += safe_append(ws_mvtec, SAMIAD_MVTEC)
    added += safe_append(ws_visa,  SAMIAD_VISA)

    print('=== Adding DiffAST (arXiv 2604.19240) ===')
    added += safe_append(ws_all,   DIFFAST_OVERVIEW)
    added += safe_append(ws_mvtec, DIFFAST_MVTEC)
    added += safe_append(ws_visa,  DIFFAST_VISA)
    added += safe_append(ws_mpdd,  DIFFAST_MPDD)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print(f'New entries added: {added} sheet rows '
          '(3 new papers: DART, SAM-IAD, DiffAST)')
    print('All 3 papers staged as ⏳ Pending — metrics not verifiable this week '
          '(arXiv + ScienceDirect blocked by proxy).')


if __name__ == '__main__':
    main()
