#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fix pass for journal harvest 2026-10-11.
Adds missing per-dataset rows and the IndusAgent overview entry.
"""

from pathlib import Path
import sys

try:
    import openpyxl
except ImportError:
    sys.path.insert(0, '/tmp/pkgs')
    import openpyxl

AD_XLSX = Path(__file__).resolve().parent.parent / 'AnomalyDetection_Papers_Summary_v10_20260425.xlsx'

# Per-dataset rows that were blocked by the global dupe-set bug in the first pass

ANCHORADAPT_MVTEC = (
    'MVTec AD (few-shot cross-dataset protocol; NON-STANDARD PROTOCOL)',
    '基於視覺語言模型 (VLM-based) / CLIP / 非對稱提示自適應 / 少樣本 (Few-Shot)',
    'Anchor and Adapt (Asymmetric Prompt Adaptation for Few-Shot Industrial Anomaly Detection)',
    'Mengyang Zhao, Teng Fu, Haiyang Yu, Ke Niu, Bin Li, Xiangyang Xue',
    'arXiv 2610.07016 (2026-10-04)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — few-shot cross-dataset protocol. arXiv HTML blocked; metrics not accessible. '
    '[基於VLM/CLIP / 非對稱提示自適應 / 少樣本]',
    'https://arxiv.org/abs/2610.07016',
    None,
)

ANCHORADAPT_VISA = (
    'VisA (few-shot cross-dataset protocol; NON-STANDARD PROTOCOL)',
    '基於視覺語言模型 (VLM-based) / CLIP / 非對稱提示自適應 / 少樣本 (Few-Shot)',
    'Anchor and Adapt (Asymmetric Prompt Adaptation for Few-Shot Industrial Anomaly Detection)',
    'Mengyang Zhao, Teng Fu, Haiyang Yu, Ke Niu, Bin Li, Xiangyang Xue',
    'arXiv 2610.07016 (2026-10-04)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — few-shot cross-dataset protocol. arXiv HTML blocked; metrics not accessible. '
    '[基於VLM/CLIP / 非對稱提示自適應 / 少樣本]',
    'https://arxiv.org/abs/2610.07016',
    None,
)

VAND4_AD2 = (
    'MVTec AD 2 (VAND 4.0 Challenge; pixel-level SegF1 metric — NOT AUROC)',
    '基準/挑戰賽報告 (Benchmark / Challenge Report)',
    'VAND 4.0 Challenge Report (Detecting Defects that Matter: Application-Driven Benchmark)',
    'MVTec, TUM, Amazon, Intel, NVIDIA (VAND 4.0 organizing committee)',
    'CVPR 2026 VAND 4.0 Workshop / arXiv 2610.04392 (2026-10-03)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — benchmark/challenge report. MVTec AD 2 Industrial Track. '
    'Primary metric: pixel-level SegF1 (NOT AUROC). Best regular-setting entry: ~57% SegF1. '
    'Zero-shot entries trail by ~15 SegF1 pts. DINOv3 backbones dominate. '
    'Full leaderboard on benchmark.mvtec.com. NOT a standalone method paper. '
    '[基準/挑戰賽報告]',
    'https://arxiv.org/abs/2610.04392',
    'https://github.com/cvpr-vand/vand-2026',
)

INDUSAGENT_OVERVIEW = (
    'MVTec AD, VisA, MPDD, DTD, SDD (零樣本開放詞彙協議; NON-STANDARD PROTOCOL — open-vocabulary zero-shot)',
    '基於基礎模型 (Foundation Model-based) / VLM / 代理工具增強 (Agentic Tools) / 開放詞彙 (Open-Vocabulary) / 零樣本',
    'IndusAgent (Reinforcing Open-Vocabulary Industrial Anomaly Detection with Agentic Tools)',
    'Rongbin Tan, Huawei Cao, Shuhan Song et al. (13 authors)',
    'arXiv 2605.20682 (2026-05-20)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2605.20682 (May 20, 2026); arXiv HTML blocked, no GitHub found. '
    'Tool-augmented agentic framework for open-vocabulary IAD. Fine-tuned on custom Indus-CoT dataset '
    'with gated reinforcement learning objective (classification + localization + anomaly reasoning + '
    'tool usage). Uses dynamic region cropping, high-frequency feature enhancement, prior retrieval. '
    'Evaluates: MVTec-AD, VisA, MPDD, DTD, SDD — zero-shot protocol. '
    'Headline: avg score 83.4% (metric unclear from web snippet; NOT verified). '
    'NON-STANDARD PROTOCOL (open-vocabulary zero-shot). '
    'Previously missed in Aug/Sep harvests. Journal harvest 2026-10-11. '
    '[基於基礎模型/VLM / 代理工具增強 / 開放詞彙 / 零樣本]',
    'https://arxiv.org/abs/2605.20682',
    None,
)

INDUSAGENT_VISA = (
    'VisA (零樣本開放詞彙協議; NON-STANDARD PROTOCOL — open-vocabulary zero-shot)',
    '基於基礎模型 (Foundation Model-based) / VLM / 代理工具增強 (Agentic Tools) / 開放詞彙 / 零樣本',
    'IndusAgent (Reinforcing Open-Vocabulary Industrial Anomaly Detection with Agentic Tools)',
    'Rongbin Tan, Huawei Cao, Shuhan Song et al.',
    'arXiv 2605.20682 (2026-05-20)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — open-vocabulary zero-shot protocol. arXiv HTML blocked; metrics not accessible. '
    '[基於基礎模型/VLM / 代理 / 零樣本]',
    'https://arxiv.org/abs/2605.20682',
    None,
)

INDUSAGENT_MPDD = (
    'MPDD (零樣本開放詞彙協議; NON-STANDARD PROTOCOL — open-vocabulary zero-shot)',
    '基於基礎模型 (Foundation Model-based) / VLM / 代理工具增強 (Agentic Tools) / 開放詞彙 / 零樣本',
    'IndusAgent (Reinforcing Open-Vocabulary Industrial Anomaly Detection with Agentic Tools)',
    'Rongbin Tan, Huawei Cao, Shuhan Song et al.',
    'arXiv 2605.20682 (2026-05-20)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — open-vocabulary zero-shot protocol. arXiv HTML blocked; metrics not accessible. '
    '[基於基礎模型/VLM / 代理 / 零樣本]',
    'https://arxiv.org/abs/2605.20682',
    None,
)


def sheet_has_link(ws, link):
    link_lower = link.lower()
    for row in ws.iter_rows(values_only=True):
        if len(row) > 13 and row[13] and str(row[13]).lower() == link_lower:
            return True
    return False


def safe_append(ws, row_tuple):
    link = row_tuple[13]
    if link and sheet_has_link(ws, link):
        print(f'  SKIP (dupe): {row_tuple[2][:55]} → {ws.title}')
        return False
    ws.append(row_tuple)
    print(f'  ADD: {row_tuple[2][:55]} → {ws.title}')
    return True


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))
    ws_all   = wb['總覽 All Papers']
    ws_mvtec = wb['MVTec AD']
    ws_visa  = wb['VisA']
    ws_ad2   = wb['MVTec AD2']
    ws_mpdd  = wb['MPDD']

    print('=== Anchor and Adapt → MVTec AD ===')
    safe_append(ws_mvtec, ANCHORADAPT_MVTEC)

    print('=== Anchor and Adapt → VisA ===')
    safe_append(ws_visa, ANCHORADAPT_VISA)

    print('=== VAND 4.0 → MVTec AD2 ===')
    safe_append(ws_ad2, VAND4_AD2)

    print('=== IndusAgent → 總覽 All Papers ===')
    safe_append(ws_all, INDUSAGENT_OVERVIEW)

    print('=== IndusAgent → VisA ===')
    safe_append(ws_visa, INDUSAGENT_VISA)

    print('=== IndusAgent → MPDD ===')
    safe_append(ws_mpdd, INDUSAGENT_MPDD)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')


if __name__ == '__main__':
    main()
