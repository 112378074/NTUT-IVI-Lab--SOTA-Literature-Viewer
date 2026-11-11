#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Weekly IVAD Journal Harvest — 2026-11-11
Coverage window: 2026-11-09 to 2026-11-11

Target journals searched:
  IEEE TPAMI, IEEE TII, IEEE TNNLS, IEEE TIP, IEEE TCSVT,
  Pattern Recognition, IJCV, Computers in Industry,
  Information Fusion, Knowledge-Based Systems,
  Information Sciences, Neural Networks

Network access this run:
  - arXiv.org: BLOCKED by environment proxy (ENOTFOUND — consistent with prior weeks)
  - IEEE Xplore: BLOCKED by environment proxy
  - ScienceDirect: BLOCKED by environment proxy
  - GitHub: ACCESSIBLE (used for VerdictAD metric verification)
  - Web search (Google snippets): ACCESSIBLE

Dupe-check results (against workbook URL column):
  - DART (2611.08271): ALREADY IN WORKBOOK (harvest 2026-11-08)
  - SAM-IAD (KBS): ALREADY IN WORKBOOK (harvest 2026-11-08)
  - DiffAST (2604.19240): ALREADY IN WORKBOOK (harvest 2026-11-08)
  - Agentic Ensemble (2610.28560): ALREADY IN WORKBOOK (prior harvest)
  - INP-Former reproduction (2611.01230): reproduction study — skip
  - TIRAD (2611.12023): NOT IN WORKBOOK — NEW this week
  - FATE (2611.08140): NOT IN WORKBOOK — NEW this week
  - VerdictAD (2611.02654): NOT IN WORKBOOK — NEW this week
  - PRISM/untrained-NN (2611.08136): NOT IN WORKBOOK — NEW this week

New entries to add this week: 4 papers

1. FATE (arXiv 2611.08140, Nov 8 2026 v2)
   "Fault-Guided Cross-Layer Prediction for Unsupervised Anomaly Detection"
   (PDF/HTML title: "FATE: Fault-Aware Transfer Error for Unsupervised Anomaly Detection")
   – Method: frozen DINOv2 ViT-L block-to-block transition failure scoring; identifies
     four sharp "fault lines" (depth segment boundaries) where adjacent-layer cosine
     similarity drops; directed predictor scores transfer failure as anomaly.
   – Datasets: MVTec AD, VisA, Real-IAD. Single-class unsupervised protocol.
   – arXiv HTML blocked; no GitHub repo found in search. NOT ranked.
   – Staged ⏳ Pending.

2. TIRAD (arXiv 2611.12023, Nov 2026)
   "TIRAD: Text-guided Image Residual Learning for Generalist Anomaly Detection"
   – Method: Text-guided residual learning for generalist (cross-domain) anomaly detection.
     Uses textual descriptions to guide image residual computation and anomaly scoring.
   – Datasets: MVTec AD, VisA, BraTS (cross-dataset transfer settings).
   – arXiv HTML blocked; no GitHub found in search. NOT ranked.
   – Staged ⏳ Pending.

3. VerdictAD (arXiv 2611.02654, Nov 2026)
   "VerdictAD: Few-Shot Anomaly Detection via Decision-Guided Refinement"
   – Authors: Shangquan Sun et al.
   – Method: Training-free few-shot framework. A visual decision module refines anomaly
     scores from an existing detector (DINOv2-L, PatchCore, AnomalyDINO, SubspaceAD)
     via decision-guided score redistribution. No training on target data required.
   – Datasets: MVTec AD, VisA, MVTec LOCO AD (1-shot, 2-shot, 4-shot settings).
   – GitHub: https://github.com/sgstq0811/VerdictAD-Jev-AD
   – VERIFIED from official GitHub README 4-shot table (mean over seeds 42, 43, 45;
     best single-detector: DINOv2-L + VerdictAD):
       MVTec AD:     I-AUROC 98.7 / P-AUROC 97.9 / AUPRO 95.7
       VisA:         I-AUROC 96.9 / P-AUROC 98.7 / AUPRO 94.4
       MVTec LOCO AD:I-AUROC 77.5 / P-AUROC 80.0 / AUPRO 65.4
   – Protocol: Few-shot AD (NOT standard unsupervised AD). Rows labeled 🟢 Few-shot AD.
   – NOTE: not directly comparable to standard unsupervised leaderboard results.

4. PRISM (arXiv 2611.08136, Nov 2026)
   "How Far Can Untrained Neural Networks Go in Anomaly Detection? From Tables to Pixels"
   – Method: Untrained neural networks as random projection; anomaly scored by
     Mahalanobis distance on randomly mapped features. Extends table-level anomaly
     detection to pixel-level localization without any training.
   – Datasets: MVTec AD, VisA (27 visual product categories).
   – arXiv HTML blocked; no GitHub found in search. NOT ranked.
   – Staged ⏳ Pending.

No SMTP credentials found (scripts/.env absent) → email skipped.
"""

from pathlib import Path
import sys

try:
    import openpyxl
except ImportError:
    sys.path.insert(0, '/usr/local/lib/python3.13/site-packages')
    import openpyxl

AD_XLSX = Path(__file__).resolve().parent.parent / 'AnomalyDetection_Papers_Summary_v10_20260425.xlsx'

# ---------------------------------------------------------------------------
# Row format (15 cols):
# (資料集, 類別, 方法, 作者, 發表, 年月, 狀態,
#  I-AUROC, P-AUROC, P-AP, P-PRO, FPS, 備註(特色/based), 連結, GitHub)
# ---------------------------------------------------------------------------

# ===========================================================================
# 1. FATE (arXiv 2611.08140, Nov 2026)
#    "Fault-Guided Cross-Layer Prediction for Unsupervised Anomaly Detection"
#    (also titled: "FATE: Fault-Aware Transfer Error for Unsupervised Anomaly Detection")
# ===========================================================================

FATE_OVERVIEW = (
    'MVTec AD, VisA, Real-IAD (Unsupervised AD)',
    '基礎模型 (Foundation Model-based) / DINOv2-based '
    '/ 表徵學習 (Representation-based) / 跨層預測 (Cross-Layer Prediction-based)',
    'FATE (Fault-Aware Transfer Error / Fault-Guided Cross-Layer Prediction '
    'for Unsupervised Anomaly Detection)',
    'Tao Wang et al.',
    'arXiv 2611.08140 (2026-11, v2 2026-11-08)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08140 (Nov 2026, v2 2026-11-08); arXiv HTML blocked; no GitHub found. '
    'Method: Exploits frozen DINOv2 ViT-L block-to-block transition failures as anomaly scores. '
    'Identifies four sharp "fault lines" (depth segment boundaries) in DINOv2 ViT-L where '
    'adjacent-layer cosine similarity drops sharply. A directed predictor scores how well deep '
    'or shallow blocks transfer information to sensitive boundary blocks; transfer failure = '
    'anomaly signal. Single-class unsupervised protocol. Evaluates on MVTec AD, VisA, Real-IAD. '
    'Journal harvest 2026-11-11. [Foundation Model / DINOv2 / Cross-Layer Prediction / '
    'Representation-based]',
    'https://arxiv.org/abs/2611.08140',
    None,
)

FATE_MVTEC = (
    'MVTec AD',
    '基礎模型 (Foundation Model-based) / DINOv2-based / 跨層預測 (Cross-Layer Prediction-based)',
    'FATE (Fault-Aware Transfer Error / Fault-Guided Cross-Layer Prediction)',
    'Tao Wang et al.',
    'arXiv 2611.08140 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08140 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports category-averaged metrics on MVTec AD. '
    '[Foundation Model / DINOv2 / Cross-Layer Prediction]',
    'https://arxiv.org/abs/2611.08140',
    None,
)

FATE_VISA = (
    'VisA',
    '基礎模型 (Foundation Model-based) / DINOv2-based / 跨層預測 (Cross-Layer Prediction-based)',
    'FATE (Fault-Aware Transfer Error / Fault-Guided Cross-Layer Prediction)',
    'Tao Wang et al.',
    'arXiv 2611.08140 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08140 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports category-averaged metrics on VisA. '
    '[Foundation Model / DINOv2 / Cross-Layer Prediction]',
    'https://arxiv.org/abs/2611.08140',
    None,
)

# ===========================================================================
# 2. TIRAD (arXiv 2611.12023, Nov 2026)
#    "TIRAD: Text-guided Image Residual Learning for Generalist Anomaly Detection"
# ===========================================================================

TIRAD_OVERVIEW = (
    'MVTec AD, VisA, BraTS (Generalist AD / cross-domain)',
    '視覺語言模型 (VLM-based) / 文字引導 (Text-guided) '
    '/ 通用型 (Generalist AD) / 殘差學習 (Residual Learning-based)',
    'TIRAD (Text-guided Image Residual Learning for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.12023)',
    'arXiv 2611.12023 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.12023 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Method: Text-guided residual learning for generalist (cross-domain) anomaly detection. '
    'Uses textual descriptions to guide image residual computation and anomaly scoring. '
    'Evaluates on MVTec AD, VisA, and BraTS (cross-dataset transfer settings). '
    'VLM/text-guided generalist AD approach. '
    'Journal harvest 2026-11-11. [VLM / Text-guided / Generalist AD / Residual Learning]',
    'https://arxiv.org/abs/2611.12023',
    None,
)

TIRAD_MVTEC = (
    'MVTec AD',
    '視覺語言模型 (VLM-based) / 文字引導 (Text-guided) / 通用型 (Generalist AD)',
    'TIRAD (Text-guided Image Residual Learning for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.12023)',
    'arXiv 2611.12023 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.12023 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports metrics on MVTec AD in cross-dataset generalist transfer setting. '
    '[VLM / Text-guided / Generalist AD]',
    'https://arxiv.org/abs/2611.12023',
    None,
)

TIRAD_VISA = (
    'VisA',
    '視覺語言模型 (VLM-based) / 文字引導 (Text-guided) / 通用型 (Generalist AD)',
    'TIRAD (Text-guided Image Residual Learning for Generalist Anomaly Detection)',
    '(authors not confirmed — arXiv 2611.12023)',
    'arXiv 2611.12023 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.12023 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports metrics on VisA in cross-dataset generalist transfer setting. '
    '[VLM / Text-guided / Generalist AD]',
    'https://arxiv.org/abs/2611.12023',
    None,
)

# ===========================================================================
# 3. VerdictAD (arXiv 2611.02654, Nov 2026)
#    "VerdictAD: Few-Shot Anomaly Detection via Decision-Guided Refinement"
#    VERIFIED from official GitHub README (sgstq0811/VerdictAD-Jev-AD)
#    4-shot table, DINOv2-L + VerdictAD, mean over seeds 42, 43, 45
#    Protocol: Few-shot AD (NOT standard unsupervised AD)
# ===========================================================================

VERDICT_SOURCE = (
    '🟢 Few-shot AD verified — 4-shot I-AUROC/P-AUROC/AUPRO from official GitHub README '
    '(sgstq0811/VerdictAD-Jev-AD), DINOv2-L + VerdictAD configuration, '
    'mean over seeds 42, 43, 45. Protocol: Few-shot AD — NOT comparable to '
    'standard unsupervised AD. '
    'Journal harvest 2026-11-11. [Representation / Training-Free / Decision-Guided / Few-shot AD]'
)

VERDICT_OVERVIEW = (
    'MVTec AD, VisA, MVTec LOCO AD (Few-shot AD; 1/2/4-shot settings)',
    '表徵學習 (Representation-based) / 訓練免 (Training-Free) '
    '/ 決策引導細化 (Decision-Guided Refinement) / 少樣本 (Few-shot AD)',
    'VerdictAD (Few-Shot Anomaly Detection via Decision-Guided Refinement)',
    'Shangquan Sun et al.',
    'arXiv 2611.02654 (2026-11)',
    '2026-11',
    '🟢 Few-shot AD verified (4-shot)',
    98.7, 97.9, None, 95.7, None,
    '🟢 Few-shot AD verified — 4-shot metrics on MVTec AD from official GitHub README '
    '(sgstq0811/VerdictAD-Jev-AD), DINOv2-L + VerdictAD, mean seeds 42,43,45. '
    'Full table: MVTec AD I-AUROC/P-AUROC/AUPRO: 98.7/97.9/95.7; '
    'VisA: 96.9/98.7/94.4; MVTec LOCO: 77.5/80.0/65.4. '
    'Training-free framework: visual decision module refines existing detector '
    '(DINOv2-L, PatchCore, AnomalyDINO, SubspaceAD) scores via decision-guided '
    'score redistribution. No training on target data required. '
    'Protocol: Few-shot AD — NOT comparable to standard unsupervised AD results. '
    'Journal harvest 2026-11-11. [Representation / Training-Free / Decision-Guided / Few-shot AD]',
    'https://arxiv.org/abs/2611.02654',
    'https://github.com/sgstq0811/VerdictAD-Jev-AD',
)

VERDICT_MVTEC = (
    'MVTec AD',
    '表徵學習 (Representation-based) / 訓練免 (Training-Free) '
    '/ 決策引導細化 (Decision-Guided Refinement) / 少樣本 (Few-shot AD)',
    'VerdictAD (Few-Shot Anomaly Detection via Decision-Guided Refinement)',
    'Shangquan Sun et al.',
    'arXiv 2611.02654 (2026-11)',
    '2026-11',
    '🟢 Few-shot AD verified (4-shot)',
    98.7, 97.9, None, 95.7, None,
    VERDICT_SOURCE + ' MVTec AD 4-shot (DINOv2-L + VerdictAD): I-AUROC 98.7 / '
    'P-AUROC 97.9 / AUPRO 95.7.',
    'https://arxiv.org/abs/2611.02654',
    'https://github.com/sgstq0811/VerdictAD-Jev-AD',
)

VERDICT_VISA = (
    'VisA',
    '表徵學習 (Representation-based) / 訓練免 (Training-Free) '
    '/ 決策引導細化 (Decision-Guided Refinement) / 少樣本 (Few-shot AD)',
    'VerdictAD (Few-Shot Anomaly Detection via Decision-Guided Refinement)',
    'Shangquan Sun et al.',
    'arXiv 2611.02654 (2026-11)',
    '2026-11',
    '🟢 Few-shot AD verified (4-shot)',
    96.9, 98.7, None, 94.4, None,
    VERDICT_SOURCE + ' VisA 4-shot (DINOv2-L + VerdictAD): I-AUROC 96.9 / '
    'P-AUROC 98.7 / AUPRO 94.4.',
    'https://arxiv.org/abs/2611.02654',
    'https://github.com/sgstq0811/VerdictAD-Jev-AD',
)

VERDICT_LOCO = (
    'MVTec LOCO',
    '表徵學習 (Representation-based) / 訓練免 (Training-Free) '
    '/ 決策引導細化 (Decision-Guided Refinement) / 少樣本 (Few-shot AD)',
    'VerdictAD (Few-Shot Anomaly Detection via Decision-Guided Refinement)',
    'Shangquan Sun et al.',
    'arXiv 2611.02654 (2026-11)',
    '2026-11',
    '🟢 Few-shot AD verified (4-shot)',
    77.5, 80.0, None, 65.4, None,
    VERDICT_SOURCE + ' MVTec LOCO AD 4-shot (DINOv2-L + VerdictAD): I-AUROC 77.5 / '
    'P-AUROC 80.0 / AUPRO 65.4.',
    'https://arxiv.org/abs/2611.02654',
    'https://github.com/sgstq0811/VerdictAD-Jev-AD',
)

# ===========================================================================
# 4. PRISM (arXiv 2611.08136, Nov 2026)
#    "How Far Can Untrained Neural Networks Go in Anomaly Detection?
#     From Tables to Pixels"
# ===========================================================================

PRISM_OVERVIEW = (
    'MVTec AD, VisA (27 visual categories; Unsupervised AD)',
    '無監督 (Unsupervised) / 未訓練神經網路 (Untrained NN-based) '
    '/ 馬氏距離 (Mahalanobis-based) / 無參數 (Parameter-free)',
    'PRISM (How Far Can Untrained Neural Networks Go in Anomaly Detection? '
    'From Tables to Pixels)',
    '(authors not confirmed — arXiv 2611.08136)',
    'arXiv 2611.08136 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08136 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Method: Untrained neural networks as random projectors; anomaly scored by '
    'Mahalanobis distance computed on randomly mapped (un-trained) features. Extends '
    'table-level anomaly detection to pixel-level localization without any training. '
    'Evaluates on MVTec AD and VisA (27 visual product categories). '
    'Journal harvest 2026-11-11. [Untrained NN / Mahalanobis / Parameter-free / Unsupervised AD]',
    'https://arxiv.org/abs/2611.08136',
    None,
)

PRISM_MVTEC = (
    'MVTec AD',
    '未訓練神經網路 (Untrained NN-based) / 馬氏距離 (Mahalanobis-based) / 無參數 (Parameter-free)',
    'PRISM (How Far Can Untrained Neural Networks Go in Anomaly Detection? '
    'From Tables to Pixels)',
    '(authors not confirmed — arXiv 2611.08136)',
    'arXiv 2611.08136 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08136 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports category-averaged metrics on MVTec AD (27 visual categories). '
    '[Untrained NN / Mahalanobis / Parameter-free]',
    'https://arxiv.org/abs/2611.08136',
    None,
)

PRISM_VISA = (
    'VisA',
    '未訓練神經網路 (Untrained NN-based) / 馬氏距離 (Mahalanobis-based) / 無參數 (Parameter-free)',
    'PRISM (How Far Can Untrained Neural Networks Go in Anomaly Detection? '
    'From Tables to Pixels)',
    '(authors not confirmed — arXiv 2611.08136)',
    'arXiv 2611.08136 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.08136 (Nov 2026); arXiv HTML blocked; no GitHub found. '
    'Reports category-averaged metrics on VisA (27 visual categories). '
    '[Untrained NN / Mahalanobis / Parameter-free]',
    'https://arxiv.org/abs/2611.08136',
    None,
)


# ---------------------------------------------------------------------------
# Insert helpers
# ---------------------------------------------------------------------------

def sheet_has_link_and_method(ws, link, method_substr):
    """Dupe check by URL (col 14) + method name substring (col 3)."""
    link_lower = link.lower()
    method_lower = method_substr.lower()
    for row in ws.iter_rows(values_only=True):
        if len(row) > 13 and row[13] and str(row[13]).lower() == link_lower:
            if len(row) > 2 and row[2] and method_lower in str(row[2]).lower():
                return True
    return False


def safe_append(ws, row_tuple):
    link = row_tuple[13] or ''
    method = str(row_tuple[2])[:30]
    if link and sheet_has_link_and_method(ws, link, method[:20]):
        print(f'  SKIP (dupe): {str(row_tuple[2])[:55]} → {ws.title}')
        return False
    ws.append(row_tuple)
    print(f'  ADD: {str(row_tuple[2])[:55]} → {ws.title}')
    return True


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))

    ws_all   = None
    ws_mvtec = wb['MVTec AD']
    ws_visa  = wb['VisA']
    ws_loco  = wb['MVTec LOCO']

    for name in wb.sheetnames:
        if 'All' in name or '總' in name or '総' in name:
            ws_all = wb[name]
            break
    if ws_all is None:
        ws_all = wb.active

    added = 0

    print('=== Adding FATE (arXiv 2611.08140) ===')
    added += safe_append(ws_all,   FATE_OVERVIEW)
    added += safe_append(ws_mvtec, FATE_MVTEC)
    added += safe_append(ws_visa,  FATE_VISA)

    print('=== Adding TIRAD (arXiv 2611.12023) ===')
    added += safe_append(ws_all,   TIRAD_OVERVIEW)
    added += safe_append(ws_mvtec, TIRAD_MVTEC)
    added += safe_append(ws_visa,  TIRAD_VISA)

    print('=== Adding VerdictAD (arXiv 2611.02654) [verified 4-shot] ===')
    added += safe_append(ws_all,   VERDICT_OVERVIEW)
    added += safe_append(ws_mvtec, VERDICT_MVTEC)
    added += safe_append(ws_visa,  VERDICT_VISA)
    added += safe_append(ws_loco,  VERDICT_LOCO)

    print('=== Adding PRISM (arXiv 2611.08136) ===')
    added += safe_append(ws_all,   PRISM_OVERVIEW)
    added += safe_append(ws_mvtec, PRISM_MVTEC)
    added += safe_append(ws_visa,  PRISM_VISA)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print(f'New entries added: {added} sheet rows (4 papers: FATE, TIRAD, VerdictAD, PRISM)')
    print()
    print('Summary:')
    print('  FATE (2611.08140):     ⏳ Pending — arXiv HTML + GitHub blocked')
    print('  TIRAD (2611.12023):    ⏳ Pending — arXiv HTML blocked, no GitHub')
    print('  VerdictAD (2611.02654):🟢 Few-shot AD verified (4-shot, GitHub README)')
    print('  PRISM (2611.08136):    ⏳ Pending — arXiv HTML blocked, no GitHub')


if __name__ == '__main__':
    main()
