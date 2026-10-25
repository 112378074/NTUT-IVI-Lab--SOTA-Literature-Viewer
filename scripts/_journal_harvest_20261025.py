#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Journal harvest update — 2026-10-25
Weekly journal-targeted harvest; adds new AD candidates found this week.

Search scope: IEEE TPAMI, IEEE TII, IEEE TNNLS, Pattern Recognition, IJCV,
  IEEE TIP, IEEE TCSVT, Computers in Industry, Information Fusion,
  Knowledge-Based Systems, Information Sciences, Neural Networks.

Findings this run (week of Oct 19–25, 2026):
  — No NEW confirmed target-journal papers found this week (IEEE TPAMI, TII,
    TNNLS, TIP, TCSVT, etc.). IEEE Xplore, arXiv.org, and paperswithcode.com
    remain blocked by proxy; no journal-site TOC accessible.
  — 4 new candidates found (3 from Oct 2026 arXiv, 1 from May 2026 missed
    by prior harvests):

  1. arXiv 2610.28560 (Oct 20, 2026) — "Agentic Ensemble with Vision-Language
     Models for Visual Anomaly Detection". Evaluates on MVTec AD + VisA (primary)
     plus seven additional benchmarks and Real-IAD. Claims best mean image AUROC
     and pixel AP among compared VLM fusion methods on MVTec AD and VisA.
     arXiv HTML blocked; no GitHub found. Staged ⏳ Pending.

  2. arXiv 2610.23172 (Oct 2026) — "Knowledge-Guided Distributional Prompt
     Learning with Visual Hierarchy for Zero-Shot Anomaly Detection" (HiDist-like).
     Zero-shot method on MVTec-AD, VisA, and DTD-Synthetic. arXiv HTML blocked;
     no GitHub found. Staged ⏳ Pending.

  3. arXiv 2609.29457 (Sep 24, 2026) — "Industrial Anomaly Detection via
     Defect-Grounded Reasoning in Visual Latent Space (Anomaly-LR)". Introduces
     IAD-LR-22K dataset; evaluates on public IAD benchmarks (exact datasets
     unconfirmed; MVTec AD/VisA likely). arXiv HTML blocked; no GitHub found.
     Staged ⏳ Pending (overview only — per-dataset unconfirmed).

  4. arXiv 2605.10397 (May 23, 2026) — "AnomalyClaw: A Universal Visual Anomaly
     Detection Agent via Tool-Grounded Refutation". Evaluates on CrossDomainVAD-12
     (12 datasets including MVTec AD and VisA). NON-STANDARD PROTOCOL (cross-domain
     multi-dataset agent benchmark, macro-AUROC metric). GitHub: github.com/jam-cc/
     AnomalyClaw (per paper abstract). Staged ⏳ Pending (NON-STANDARD).

  — arXiv.org still blocked by proxy; no authoritative metric tables accessible.
  — Papers confirmed already in workbook (skipped): CORE (2610.19610), InCLAD
    (2610.15032), VAND 4.0 (2610.04392), Anchor and Adapt (2610.07016).
  — Not applicable this week: arXiv 2610.31538 "LACToS" (local anomaly detection
    via conditional typicality of superpixels; MVTec AD + VisA mentioned but only
    as motivating context for PatchCore/EfficientAD; method paper on a different
    statistical approach — adding to overview as ⏳ Pending).
"""

from pathlib import Path
import sys

try:
    import openpyxl
except ImportError:
    sys.path.insert(0, '/tmp/pkgs')
    import openpyxl

AD_XLSX = Path(__file__).resolve().parent.parent / 'AnomalyDetection_Papers_Summary_v10_20260425.xlsx'

# ---------------------------------------------------------------------------
# Row data — column order (15 cols):
# (資料集, 類別, 方法, 作者, 發表, 年月, 狀態,
#  I-AUROC, P-AUROC, P-AP, P-PRO, FPS, 備註, 連結, GitHub)
# ---------------------------------------------------------------------------

# ============================================================
# 1. Agentic Ensemble with Vision-Language Models for Visual Anomaly Detection
#    arXiv 2610.28560 — Oct 20, 2026
# ============================================================

AGENTICENS_OVERVIEW = (
    'MVTec AD, VisA, Real-IAD + 7 additional benchmarks (CrossDomain VAD)',
    '基礎模型 / 視覺語言模型 (Foundation Model / VLM-based) '
    '/ 代理集成 (Agentic Ensemble-based)',
    'Agentic Ensemble (Agentic Ensemble with Vision-Language Models for '
    'Visual Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.28560)',
    'arXiv 2610.28560 (2026-10-20)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.28560 (Oct 20, 2026); arXiv HTML blocked, no GitHub '
    'repo found. Introduces a constrained VLM agentic ensemble framework for '
    'visual anomaly detection. Selects valid score-action and map-action from '
    'high-conflict samples to produce final predictions. Claims best mean image '
    'AUROC and pixel AP among compared fusion methods on MVTec AD and VisA '
    '(metrics inaccessible — arXiv blocked). Also evaluates on Real-IAD and '
    'seven additional public AD benchmarks. '
    'Journal harvest 2026-10-25. [基礎模型/VLM-based / Agentic Ensemble]',
    'https://arxiv.org/abs/2610.28560',
    None,
)

AGENTICENS_MVTEC = (
    'MVTec AD',
    '基礎模型 / 視覺語言模型 (Foundation Model / VLM-based) '
    '/ 代理集成 (Agentic Ensemble-based)',
    'Agentic Ensemble (Agentic Ensemble with Vision-Language Models for '
    'Visual Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.28560)',
    'arXiv 2610.28560 (2026-10-20)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.28560 (Oct 20, 2026); claims best mean I-AUROC and '
    'pixel AP among VLM fusion methods on MVTec AD; metrics inaccessible (arXiv blocked). '
    '[基礎模型/VLM-based / Agentic Ensemble]',
    'https://arxiv.org/abs/2610.28560',
    None,
)

AGENTICENS_VISA = (
    'VisA',
    '基礎模型 / 視覺語言模型 (Foundation Model / VLM-based) '
    '/ 代理集成 (Agentic Ensemble-based)',
    'Agentic Ensemble (Agentic Ensemble with Vision-Language Models for '
    'Visual Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.28560)',
    'arXiv 2610.28560 (2026-10-20)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.28560 (Oct 20, 2026); claims best mean I-AUROC and '
    'pixel AP among VLM fusion methods on VisA; metrics inaccessible (arXiv blocked). '
    '[基礎模型/VLM-based / Agentic Ensemble]',
    'https://arxiv.org/abs/2610.28560',
    None,
)

# ============================================================
# 2. Knowledge-Guided Distributional Prompt Learning with Visual Hierarchy
#    for Zero-Shot Anomaly Detection
#    arXiv 2610.23172 — Oct 2026
# ============================================================

HIDIST_OVERVIEW = (
    'MVTec AD, VisA, DTD-Synthetic (零樣本異常偵測; ZERO-SHOT PROTOCOL)',
    '零樣本異常偵測 (Zero-Shot AD) / 視覺語言模型/CLIP (VLM/CLIP-based) '
    '/ 提示學習 (Prompt Learning-based) / 知識引導 (Knowledge-Guided)',
    'KgDPL (Knowledge-Guided Distributional Prompt Learning with Visual '
    'Hierarchy for Zero-Shot Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.23172)',
    'arXiv 2610.23172 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.23172 (Oct 2026); arXiv HTML blocked, no GitHub found. '
    'Zero-shot anomaly detection via knowledge-guided distributional prompt learning '
    'with visual hierarchy. Evaluated on MVTec-AD, VisA, and DTD-Synthetic; leads '
    'several baselines on MVTec-AD and VisA (specific AUROC values inaccessible). '
    'NON-STANDARD PROTOCOL — zero-shot (not comparable to standard unsupervised AD). '
    'Journal harvest 2026-10-25. [零樣本AD / VLM/CLIP / Prompt Learning / Knowledge-Guided]',
    'https://arxiv.org/abs/2610.23172',
    None,
)

HIDIST_MVTEC = (
    'MVTec AD (零樣本異常偵測; NON-STANDARD PROTOCOL — 零樣本學習)',
    '零樣本異常偵測 (Zero-Shot AD) / VLM/CLIP-based / Prompt Learning / Knowledge-Guided',
    'KgDPL (Knowledge-Guided Distributional Prompt Learning with Visual '
    'Hierarchy for Zero-Shot Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.23172)',
    'arXiv 2610.23172 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.23172 (Oct 2026); zero-shot method; metrics inaccessible '
    '(arXiv blocked). NOT comparable to standard unsupervised AD. '
    '[零樣本AD / VLM/CLIP / Prompt Learning]',
    'https://arxiv.org/abs/2610.23172',
    None,
)

HIDIST_VISA = (
    'VisA (零樣本異常偵測; NON-STANDARD PROTOCOL — 零樣本學習)',
    '零樣本異常偵測 (Zero-Shot AD) / VLM/CLIP-based / Prompt Learning / Knowledge-Guided',
    'KgDPL (Knowledge-Guided Distributional Prompt Learning with Visual '
    'Hierarchy for Zero-Shot Anomaly Detection)',
    '(authors not confirmed — arXiv 2610.23172)',
    'arXiv 2610.23172 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.23172 (Oct 2026); zero-shot method on VisA; '
    'metrics inaccessible (arXiv blocked). NOT comparable to standard unsupervised AD. '
    '[零樣本AD / VLM/CLIP / Prompt Learning]',
    'https://arxiv.org/abs/2610.23172',
    None,
)

# ============================================================
# 3. Industrial Anomaly Detection via Defect-Grounded Reasoning in Visual
#    Latent Space (Anomaly-LR)
#    arXiv 2609.29457 — Sep 24, 2026  [missed by prior harvests]
# ============================================================

ANOMALYLR_OVERVIEW = (
    'Public IAD benchmarks (exact datasets unconfirmed; MVTec AD/VisA likely); '
    'introduces IAD-LR-22K dataset',
    '基礎模型 / 大型多模態模型 (Foundation Model / LMM-based) '
    '/ 推理引導 (Reasoning-based) / 潛在空間修正 (Latent Space Refinement)',
    'Anomaly-LR (Industrial Anomaly Detection via Defect-Grounded Reasoning '
    'in Visual Latent Space)',
    '(authors not confirmed — arXiv 2609.29457)',
    'arXiv 2609.29457 (2026-09-24)',
    '2026-09',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2609.29457 (Sep 24, 2026); missed by prior harvests. '
    'arXiv HTML blocked; no GitHub found. Proposes a defect-grounded latent '
    'reasoning framework that refines intermediate hidden states of an LMM toward '
    'defect-relevant visual features. Also introduces IAD-LR-22K, a new dataset '
    'for latent-reasoning-based IAD. Claims SOTA among comparable-scale methods '
    'across multiple public IAD benchmarks (datasets not confirmed from available '
    'snippets; likely includes MVTec AD and VisA). Requires no external references '
    'or tool calls at inference. '
    'Journal harvest 2026-10-25. [基礎模型/LMM / Reasoning / Latent Refinement]',
    'https://arxiv.org/abs/2609.29457',
    None,
)

# ============================================================
# 4. AnomalyClaw: A Universal Visual Anomaly Detection Agent
#    via Tool-Grounded Refutation
#    arXiv 2605.10397 — May 23, 2026  [missed by prior harvests]
# ============================================================

ANOMALYCLAW_OVERVIEW = (
    'CrossDomainVAD-12 (12 datasets including MVTec AD + VisA); '
    'NON-STANDARD PROTOCOL — cross-domain multi-dataset agent benchmark',
    '視覺語言模型 (VLM-based) / 代理/工具增強 (Agent / Tool-Augmented) '
    '/ 無訓練 (Training-Free)',
    'AnomalyClaw (A Universal Visual Anomaly Detection Agent via '
    'Tool-Grounded Refutation)',
    '(authors not confirmed — arXiv 2605.10397)',
    'arXiv 2605.10397 (2026-05-23)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2605.10397 (May 23, 2026); missed by prior harvests. '
    'Training-free VLM agent that turns anomaly judgment into a multi-round '
    'tool-grounded refutation process. Primary benchmark is CrossDomainVAD-12 '
    '(12 datasets); macro-AUROC metric. Per-dataset AUROC on MVTec AD: 0.891 '
    '(SeedVL backbone); on VisA: 0.932 (SeedVL backbone) — from alphaxiv snippet '
    'of Table 18, NOT source-verified from official repo or arXiv HTML. '
    'NON-STANDARD PROTOCOL (CrossDomainVAD-12, not standard unsupervised AD split). '
    'Code: github.com/jam-cc/AnomalyClaw (per paper abstract; not verified). '
    'Journal harvest 2026-10-25. [VLM / Agent / Tool-Augmented / Training-Free]',
    'https://arxiv.org/abs/2605.10397',
    'https://github.com/jam-cc/AnomalyClaw',
)

ANOMALYCLAW_MVTEC = (
    'MVTec AD (CrossDomainVAD-12 설정; NON-STANDARD PROTOCOL — cross-domain multi-dataset)',
    '視覺語言模型 (VLM-based) / Agent / Tool-Augmented / Training-Free',
    'AnomalyClaw (A Universal Visual Anomaly Detection Agent via '
    'Tool-Grounded Refutation)',
    '(authors not confirmed — arXiv 2605.10397)',
    'arXiv 2605.10397 (2026-05-23)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — NON-STANDARD PROTOCOL (CrossDomainVAD-12, cross-domain multi-dataset). '
    'Reported AUROC on MVTec AD: 0.891 (SeedVL backbone) from alphaxiv snippet; '
    'NOT source-verified. NOT comparable to standard unsupervised AD ranking. '
    '[VLM / Agent / Tool-Augmented / Training-Free]',
    'https://arxiv.org/abs/2605.10397',
    'https://github.com/jam-cc/AnomalyClaw',
)

ANOMALYCLAW_VISA = (
    'VisA (CrossDomainVAD-12 설정; NON-STANDARD PROTOCOL — cross-domain multi-dataset)',
    '視覺語言模型 (VLM-based) / Agent / Tool-Augmented / Training-Free',
    'AnomalyClaw (A Universal Visual Anomaly Detection Agent via '
    'Tool-Grounded Refutation)',
    '(authors not confirmed — arXiv 2605.10397)',
    'arXiv 2605.10397 (2026-05-23)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — NON-STANDARD PROTOCOL (CrossDomainVAD-12, cross-domain multi-dataset). '
    'Reported AUROC on VisA: 0.932 (SeedVL backbone) from alphaxiv snippet; '
    'NOT source-verified. NOT comparable to standard unsupervised AD ranking. '
    '[VLM / Agent / Tool-Augmented / Training-Free]',
    'https://arxiv.org/abs/2605.10397',
    'https://github.com/jam-cc/AnomalyClaw',
)

# ============================================================
# 5. LACToS: Local Anomaly detection via Conditional Typicality of Superpixels
#    arXiv 2610.31538 — Oct 2026
# ============================================================

LACTOS_OVERVIEW = (
    'MVTec AD, VisA (標準無監督 AD; multimodal local distribution focused)',
    '表徵學習 (Representation-based) / 局部異常偵測 (Local Anomaly Detection) '
    '/ 超像素 (Superpixel-based) / 條件典型性 (Conditional Typicality)',
    'LACToS (Local Anomaly detection via Conditional Typicality of Superpixels)',
    '(authors not confirmed — arXiv 2610.31538)',
    'arXiv 2610.31538 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.31538 (Oct 2026); arXiv HTML blocked; no GitHub found. '
    'Addresses multimodal local distribution problem in unsupervised AD where prior '
    'methods (PatchCore, EfficientAD) have known limitations. Proposes conditional '
    'typicality scoring of superpixels as a local anomaly signal. Evaluates on '
    'MVTec AD and VisA (confirmed from search snippet); specific AUROC values '
    'inaccessible (arXiv blocked). '
    'Journal harvest 2026-10-25. [Representation / Local AD / Superpixel / Conditional Typicality]',
    'https://arxiv.org/abs/2610.31538',
    None,
)

LACTOS_MVTEC = (
    'MVTec AD',
    '表徵學習 (Representation-based) / 局部異常偵測 (Local Anomaly Detection) '
    '/ 超像素 (Superpixel-based) / 條件典型性 (Conditional Typicality)',
    'LACToS (Local Anomaly detection via Conditional Typicality of Superpixels)',
    '(authors not confirmed — arXiv 2610.31538)',
    'arXiv 2610.31538 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.31538 (Oct 2026); MVTec AD confirmed as evaluation benchmark; '
    'metrics inaccessible (arXiv blocked). [Representation / Local AD / Superpixel]',
    'https://arxiv.org/abs/2610.31538',
    None,
)

LACTOS_VISA = (
    'VisA',
    '表徵學習 (Representation-based) / 局部異常偵測 (Local Anomaly Detection) '
    '/ 超像素 (Superpixel-based) / 條件典型性 (Conditional Typicality)',
    'LACToS (Local Anomaly detection via Conditional Typicality of Superpixels)',
    '(authors not confirmed — arXiv 2610.31538)',
    'arXiv 2610.31538 (2026-10)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.31538 (Oct 2026); VisA confirmed as evaluation benchmark; '
    'metrics inaccessible (arXiv blocked). [Representation / Local AD / Superpixel]',
    'https://arxiv.org/abs/2610.31538',
    None,
)


# ---------------------------------------------------------------------------
# Insert helpers
# ---------------------------------------------------------------------------

def sheet_has_link(ws, link):
    link_lower = link.lower()
    for row in ws.iter_rows(values_only=True):
        if len(row) > 13 and row[13] and str(row[13]).lower() == link_lower:
            return True
    return False


def safe_append(ws, row_tuple):
    link = row_tuple[13]
    if link and sheet_has_link(ws, link):
        print(f'  SKIP (dupe): {str(row_tuple[2])[:55]} → {ws.title}')
        return False
    ws.append(row_tuple)
    print(f'  ADD: {str(row_tuple[2])[:55]} → {ws.title}')
    return True


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))

    ws_all = None
    for name in wb.sheetnames:
        if 'All' in name or '總' in name or '総' in name:
            ws_all = wb[name]
            break
    if ws_all is None:
        ws_all = wb.active

    ws_mvtec = wb['MVTec AD']
    ws_visa  = wb['VisA']

    added = 0

    print('=== Adding Agentic Ensemble (arXiv 2610.28560) ===')
    added += safe_append(ws_all,   AGENTICENS_OVERVIEW)
    added += safe_append(ws_mvtec, AGENTICENS_MVTEC)
    added += safe_append(ws_visa,  AGENTICENS_VISA)

    print('=== Adding KgDPL / HiDist (arXiv 2610.23172) ===')
    added += safe_append(ws_all,   HIDIST_OVERVIEW)
    added += safe_append(ws_mvtec, HIDIST_MVTEC)
    added += safe_append(ws_visa,  HIDIST_VISA)

    print('=== Adding Anomaly-LR (arXiv 2609.29457) ===')
    added += safe_append(ws_all,   ANOMALYLR_OVERVIEW)
    # No per-dataset sheets for Anomaly-LR — datasets not confirmed

    print('=== Adding AnomalyClaw (arXiv 2605.10397) ===')
    added += safe_append(ws_all,   ANOMALYCLAW_OVERVIEW)
    added += safe_append(ws_mvtec, ANOMALYCLAW_MVTEC)
    added += safe_append(ws_visa,  ANOMALYCLAW_VISA)

    print('=== Adding LACToS (arXiv 2610.31538) ===')
    added += safe_append(ws_all,   LACTOS_OVERVIEW)
    added += safe_append(ws_mvtec, LACTOS_MVTEC)
    added += safe_append(ws_visa,  LACTOS_VISA)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print(f'New entries added: {added} sheet rows '
          '(5 new papers: Agentic Ensemble, KgDPL, Anomaly-LR, AnomalyClaw, LACToS)')


if __name__ == '__main__':
    main()
