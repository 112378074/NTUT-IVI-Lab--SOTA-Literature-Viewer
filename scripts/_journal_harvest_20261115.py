#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Weekly IVAD Journal Harvest — 2026-11-15
Coverage window: 2026-11-12 to 2026-11-15

Target journals searched:
  IEEE TPAMI, IEEE TII, IEEE TNNLS, IEEE TIP, IEEE TCSVT,
  Pattern Recognition, IJCV, Computers in Industry,
  Information Fusion, Knowledge-Based Systems,
  Information Sciences, Neural Networks

Network access this run:
  - arXiv.org: BLOCKED by environment proxy (consistent with prior weeks)
  - IEEE Xplore: BLOCKED by environment proxy
  - ScienceDirect: BLOCKED by environment proxy
  - GitHub: ACCESSIBLE (used for metric verification via raw.githubusercontent.com)
  - Web search (Google snippets): ACCESSIBLE

Coverage note:
  No new papers with IDs 2611.13xxx–2611.17xxx (expected Nov 12–15 range)
  were found in search results indexed by the search engine. The two papers
  added this run are ICLR 2026 / BMVC 2026 acceptances from Oct 2025 arXiv
  preprints (IDs 2510.xxxxx) that were missed by prior automated harvests.

Dupe-check results (against workbook URL column):
  - FoundAD (2510.01934):         NOT IN WORKBOOK — NEW this harvest
  - Robust-FSAD (2510.13643):     NOT IN WORKBOOK — NEW this harvest

Pending backlog retried this run:
  - FATE (2611.08140):   no new GitHub found → still ⏳ Pending
  - TIRAD (2611.12023):  no new GitHub found → still ⏳ Pending
  - PRISM (2611.08136):  no new GitHub found → still ⏳ Pending
  - DART (2611.08271):   no new GitHub found → still ⏳ Pending

New entries to add this week: 2 papers

1. FoundAD (arXiv 2510.01934, ICLR 2026)
   "Foundation Visual Encoders Are Secretly Few-Shot Anomaly Detectors"
   – Authors: Guangyao Zhai, Yue Zhou, Xinyan Deng, Lars Heckler,
     Nassir Navab, Benjamin Busam (TU Munich / MVTec Software GmbH)
   – Method: Frozen foundation visual encoder (DINOv3 ViT-B) + lightweight
     nonlinear projector onto natural image manifold. Anomaly amount tracked
     via embedding difference. No text prompts. Single unified model for
     all classes (multi-class one-model few-shot setting).
   – Datasets: MVTec AD, VisA (1/2/4-shot settings).
   – GitHub README (ymxlzgy/FoundAD): contains training/inference commands
     and pre-trained model links but NO numeric results table.
   – arXiv HTML: blocked. ICLR proceedings PDF: blocked.
   – Status: NOT RANKED. Staged ⏳ Pending.
   – Search snippets mention MVTec AD I-AUROC ~96–97%, VisA P-AUROC ~99.7%
     across shot settings, but these are from third-party reviews/aggregators
     (Moonlight, EmergentMind) — NOT from an accessible authoritative table.
     NOT used per workflow rules.

2. Adversarially Robust FSAD (arXiv 2510.13643, BMVC 2026)
   "Adversarially Robust Few-Shot Anomaly Detection with Vision Foundation Models"
   – Authors: Akib Mohammed Khan, Bartosz Krawczyk
   – Method: Evaluates adversarial robustness of patch-based few-shot AD
     (DINOv2 backbone + k-NN memory bank) under white-box FGSM and PGD
     attacks. Two defenses: PatchShift (sub-patch shift scoring) and
     FeaturePurifier (lightweight MLP trained on adversarial/clean pairs).
     Uses differentiable DistanceProbe surrogate for gradient-based attacks.
   – Datasets: MVTec AD, VisA, MPDD (adversarial robustness protocol,
     NOT standard few-shot AD benchmark comparisons).
   – GitHub README (akibkhan1/robust-fsad): contains only usage commands,
     no numeric results table.
   – arXiv HTML: blocked. BMVC 2026 paper link: TBD.
   – Status: NOT RANKED. Staged ⏳ Pending.
   – NOTE: This paper's PRIMARY contribution is adversarial robustness, not
     standard AUROC improvement. Its "clean baseline" is AnomalyDINO-style
     DINOv2 k-NN, not a new AD method. Protocol: adversarial robustness
     (NON-STANDARD — not comparable to standard few-shot or unsupervised AD).
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
# 1. FoundAD (arXiv 2510.01934, ICLR 2026)
#    "Foundation Visual Encoders Are Secretly Few-Shot Anomaly Detectors"
# ===========================================================================

FOUNDAD_OVERVIEW = (
    'MVTec AD, VisA (Few-shot AD; 1/2/4-shot settings)',
    '基礎模型 (Foundation Model-based) / DINOv3-based '
    '/ 流形投影 (Manifold Projection-based) / 少樣本 (Few-shot AD)',
    'FoundAD (Foundation Visual Encoders Are Secretly Few-Shot Anomaly Detectors)',
    'Guangyao Zhai, Yue Zhou, Xinyan Deng, Lars Heckler, Nassir Navab, '
    'Benjamin Busam (TU Munich / MVTec Software GmbH)',
    'ICLR 2026 / arXiv 2510.01934 (2025-10-02)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — ICLR 2026 / arXiv 2510.01934 (Oct 2025). '
    'GitHub README (ymxlzgy/FoundAD): contains training/inference commands and '
    'pre-trained model links but NO numeric results table; arXiv HTML blocked; '
    'ICLR proceedings PDF blocked. '
    'Method: Frozen foundation visual encoder (DINOv3 ViT-B) + lightweight nonlinear '
    'projector onto the natural image manifold. Anomaly scored by embedding difference. '
    'No text prompts. Single unified model per dataset, all classes. '
    'Multi-class few-shot setting. Evaluates on MVTec AD and VisA (1/2/4-shot). '
    'Search snippets cite DINOv3 as best encoder, P-AUROC ~99.7% on VisA across '
    'shot settings — from third-party reviews only, NOT from an accessible table. '
    'NOT ranked. Protocol: Few-shot AD — not comparable to standard unsupervised AD. '
    'Journal harvest 2026-11-15. '
    '[Foundation Model / DINOv3 / Manifold Projection / Few-shot AD]',
    'https://arxiv.org/abs/2510.01934',
    'https://github.com/ymxlzgy/FoundAD',
)

FOUNDAD_MVTEC = (
    'MVTec AD (few-shot; NON-STANDARD PROTOCOL — 少樣本 few-shot AD)',
    '基礎模型 (Foundation Model-based) / DINOv3-based / 流形投影 (Manifold Projection)',
    'FoundAD (Foundation Visual Encoders Are Secretly Few-Shot Anomaly Detectors)',
    'Guangyao Zhai, Yue Zhou, Xinyan Deng, Lars Heckler, Nassir Navab, '
    'Benjamin Busam (TU Munich / MVTec Software GmbH)',
    'ICLR 2026 / arXiv 2510.01934 (2025-10-02)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — ICLR 2026. GitHub README has no results table; arXiv HTML blocked. '
    'Reports MVTec AD results at 1/2/4-shot (DINOv3 ViT-B backbone). '
    'Protocol: Few-shot AD — NOT comparable to standard unsupervised AD results. '
    '[Foundation Model / DINOv3 / Manifold Projection / Few-shot AD]',
    'https://arxiv.org/abs/2510.01934',
    'https://github.com/ymxlzgy/FoundAD',
)

FOUNDAD_VISA = (
    'VisA (few-shot; NON-STANDARD PROTOCOL — 少樣本 few-shot AD)',
    '基礎模型 (Foundation Model-based) / DINOv3-based / 流形投影 (Manifold Projection)',
    'FoundAD (Foundation Visual Encoders Are Secretly Few-Shot Anomaly Detectors)',
    'Guangyao Zhai, Yue Zhou, Xinyan Deng, Lars Heckler, Nassir Navab, '
    'Benjamin Busam (TU Munich / MVTec Software GmbH)',
    'ICLR 2026 / arXiv 2510.01934 (2025-10-02)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — ICLR 2026. GitHub README has no results table; arXiv HTML blocked. '
    'Reports VisA results at 1/2/4-shot (DINOv3 ViT-B backbone). '
    'Protocol: Few-shot AD — NOT comparable to standard unsupervised AD results. '
    '[Foundation Model / DINOv3 / Manifold Projection / Few-shot AD]',
    'https://arxiv.org/abs/2510.01934',
    'https://github.com/ymxlzgy/FoundAD',
)

# ===========================================================================
# 2. Adversarially Robust FSAD (arXiv 2510.13643, BMVC 2026)
#    "Adversarially Robust Few-Shot Anomaly Detection with Vision Foundation Models"
# ===========================================================================

ROBUST_FSAD_OVERVIEW = (
    'MVTec AD, VisA, MPDD (Adversarial robustness protocol; NON-STANDARD)',
    '基礎模型 (Foundation Model-based) / DINOv2-based '
    '/ 對抗強健性 (Adversarial Robustness-based) / 少樣本 (Few-shot AD)',
    'Adversarially Robust FSAD (Adversarially Robust Few-Shot Anomaly Detection '
    'with Vision Foundation Models)',
    'Akib Mohammed Khan, Bartosz Krawczyk',
    'BMVC 2026 / arXiv 2510.13643 (2025-10)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — BMVC 2026 / arXiv 2510.13643 (Oct 2025). '
    'GitHub README (akibkhan1/robust-fsad): usage commands only, no results table; '
    'arXiv HTML blocked; BMVC paper link TBD. '
    'Method: Evaluates adversarial robustness of patch-based few-shot AD (DINOv2 ViT '
    'backbone + k-NN memory bank) under white-box FGSM and PGD L∞ attacks. '
    'Proposes two defenses: PatchShift (sub-patch shift scoring + spatial filtering) '
    'and FeaturePurifier (lightweight MLP trained on adversarial/clean feature pairs). '
    'Uses differentiable DistanceProbe surrogate for gradient computation. '
    'Base detector: AnomalyDINO (Damm et al., WACV 2025). '
    'Evaluates on MVTec AD (15 categories), VisA (12 categories), MPDD (6 categories). '
    'PRIMARY focus: adversarial robustness, NOT standard AUROC improvement. '
    'Clean baseline = AnomalyDINO-style k-NN, not a new AD method. '
    'Protocol: adversarial robustness — NOT comparable to standard unsupervised or '
    'few-shot AD results. '
    'Journal harvest 2026-11-15. '
    '[Foundation Model / DINOv2 / Adversarial Robustness / Few-shot AD]',
    'https://arxiv.org/abs/2510.13643',
    'https://github.com/akibkhan1/robust-fsad',
)

ROBUST_FSAD_MVTEC = (
    'MVTec AD (adversarial robustness protocol; NON-STANDARD)',
    '基礎模型 (Foundation Model-based) / DINOv2-based '
    '/ 對抗強健性 (Adversarial Robustness-based)',
    'Adversarially Robust FSAD (Adversarially Robust Few-Shot Anomaly Detection '
    'with Vision Foundation Models)',
    'Akib Mohammed Khan, Bartosz Krawczyk',
    'BMVC 2026 / arXiv 2510.13643 (2025-10)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — BMVC 2026. GitHub README has no results table; arXiv HTML blocked. '
    'Reports MVTec AD results under clean and adversarial (FGSM/PGD) conditions. '
    'Protocol: adversarial robustness — NOT comparable to standard unsupervised AD. '
    '[Foundation Model / DINOv2 / Adversarial Robustness / Few-shot AD]',
    'https://arxiv.org/abs/2510.13643',
    'https://github.com/akibkhan1/robust-fsad',
)

ROBUST_FSAD_VISA = (
    'VisA (adversarial robustness protocol; NON-STANDARD)',
    '基礎模型 (Foundation Model-based) / DINOv2-based '
    '/ 對抗強健性 (Adversarial Robustness-based)',
    'Adversarially Robust FSAD (Adversarially Robust Few-Shot Anomaly Detection '
    'with Vision Foundation Models)',
    'Akib Mohammed Khan, Bartosz Krawczyk',
    'BMVC 2026 / arXiv 2510.13643 (2025-10)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — BMVC 2026. GitHub README has no results table; arXiv HTML blocked. '
    'Reports VisA results under clean and adversarial (FGSM/PGD) conditions. '
    'Protocol: adversarial robustness — NOT comparable to standard unsupervised AD. '
    '[Foundation Model / DINOv2 / Adversarial Robustness / Few-shot AD]',
    'https://arxiv.org/abs/2510.13643',
    'https://github.com/akibkhan1/robust-fsad',
)

ROBUST_FSAD_MPDD = (
    'MPDD (adversarial robustness protocol; NON-STANDARD)',
    '基礎模型 (Foundation Model-based) / DINOv2-based '
    '/ 對抗強健性 (Adversarial Robustness-based)',
    'Adversarially Robust FSAD (Adversarially Robust Few-Shot Anomaly Detection '
    'with Vision Foundation Models)',
    'Akib Mohammed Khan, Bartosz Krawczyk',
    'BMVC 2026 / arXiv 2510.13643 (2025-10)',
    '2025-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — BMVC 2026. GitHub README has no results table; arXiv HTML blocked. '
    'Reports MPDD results under clean and adversarial (FGSM/PGD) conditions. '
    'Protocol: adversarial robustness — NOT comparable to standard unsupervised AD. '
    '[Foundation Model / DINOv2 / Adversarial Robustness / Few-shot AD]',
    'https://arxiv.org/abs/2510.13643',
    'https://github.com/akibkhan1/robust-fsad',
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
    method = str(row_tuple[2])
    if link and sheet_has_link_and_method(ws, link, method[:20]):
        print(f'  SKIP (dupe): {method[:55]} → {ws.title}')
        return False
    ws.append(row_tuple)
    print(f'  ADD: {method[:55]} → {ws.title}')
    return True


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))

    ws_all  = None
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

    print('=== Adding FoundAD (arXiv 2510.01934, ICLR 2026) ===')
    added += safe_append(ws_all,   FOUNDAD_OVERVIEW)
    added += safe_append(ws_mvtec, FOUNDAD_MVTEC)
    added += safe_append(ws_visa,  FOUNDAD_VISA)

    print('=== Adding Adversarially Robust FSAD (arXiv 2510.13643, BMVC 2026) ===')
    added += safe_append(ws_all,   ROBUST_FSAD_OVERVIEW)
    added += safe_append(ws_mvtec, ROBUST_FSAD_MVTEC)
    added += safe_append(ws_visa,  ROBUST_FSAD_VISA)
    added += safe_append(ws_mpdd,  ROBUST_FSAD_MPDD)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print(f'New entries added: {added} sheet rows (2 papers: FoundAD, Robust-FSAD)')
    print()
    print('Summary:')
    print('  FoundAD (2510.01934):         ⏳ Pending — GitHub README has no results table')
    print('  Adversarially Robust FSAD (2510.13643): ⏳ Pending — GitHub README has no results table')
    print()
    print('Pending backlog status (no changes):')
    print('  FATE (2611.08140):   ⏳ Pending — no new GitHub found')
    print('  TIRAD (2611.12023):  ⏳ Pending — no new GitHub found')
    print('  PRISM (2611.08136):  ⏳ Pending — no new GitHub found')
    print('  DART (2611.08271):   ⏳ Pending — no new GitHub found')


if __name__ == '__main__':
    main()
