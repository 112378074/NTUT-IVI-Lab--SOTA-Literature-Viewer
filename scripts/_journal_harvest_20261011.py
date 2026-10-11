#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Journal harvest update — 2026-10-11
Weekly journal-targeted harvest; adds new AD candidates found this week.

Search scope: IEEE TPAMI, IEEE TII, IEEE TNNLS, Pattern Recognition, IJCV,
  IEEE TIP, IEEE TCSVT, Computers in Industry, Information Fusion,
  Knowledge-Based Systems, Information Sciences, Neural Networks.

Findings this run:
  — 2 new October 2026 arXiv papers found (2610.07016, 2610.04392):
    * Anchor and Adapt (arXiv 2610.07016, Oct 4, 2026) — few-shot, VLM/CLIP prompt learning,
      MVTec AD + VisA. arXiv HTML blocked, no GitHub repo found. Staged ⏳ Pending.
    * VAND 4.0 Challenge report (arXiv 2610.04392, Oct 3, 2026) — benchmark/challenge
      report at CVPR 2026; MVTec AD 2 + Kaputt 2. Best unsupervised ~57% SegF1 on
      MVTec AD 2; zero-shot trails by ~15 SegF1 points; DINOv3 dominates. Staged ⏳ Pending.
  — 1 previously missed arXiv paper added (2605.20682):
    * IndusAgent (arXiv 2605.20682, May 20, 2026) — open-vocabulary zero-shot, VLM/agent-based,
      MVTec AD + VisA + MPDD + DTD + SDD. Headline avg 83.4% (metric unclear from web snippet;
      NOT verified from authoritative table). Staged ⏳ Pending.
  — arXiv.org still blocked by proxy; no GitHub repos found for any new candidates.
  — 2608.17362 (Continuity-Driven) excluded: evaluates on NEU-DET/industrial metal only,
    not MVTec AD / VisA family — not applicable for this AD workbook.
  — All 3 new entries staged as ⏳ Pending with blank metrics.
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

# --- 1. Anchor and Adapt — arXiv 2610.07016, few-shot VLM/CLIP-based ---

ANCHORADAPT_OVERVIEW = (
    'MVTec AD, VisA (few-shot cross-dataset protocol; NON-STANDARD PROTOCOL)',
    '基於視覺語言模型 (VLM-based) / CLIP / 非對稱提示自適應 (Asymmetric Prompt Adaptation) / 少樣本 (Few-Shot)',
    'Anchor and Adapt (Asymmetric Prompt Adaptation for Few-Shot Industrial Anomaly Detection)',
    'Mengyang Zhao, Teng Fu, Haiyang Yu, Ke Niu, Bin Li, Xiangyang Xue',
    'arXiv 2610.07016 (2026-10-04)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.07016 (Oct 4, 2026); arXiv HTML blocked, no GitHub repo found. '
    'Two-stage prompt learning framework: Stage I learns transferable normal/abnormal text anchors '
    'from annotated auxiliary data; Stage II freezes anchors and adapts an additional normal branch '
    'using few target normal images. Text-anchor regularization keeps adapted prompts consistent '
    'with generic normal prior and separated from abnormal anchors. Evaluates cross-dataset transfer '
    'in both directions between MVTec AD and VisA under 1-, 2-shot settings. Competes with PromptAD '
    'and template-based prompt methods. NON-STANDARD PROTOCOL (few-shot). '
    'Journal harvest 2026-10-11. [基於VLM/CLIP / 非對稱提示自適應 / 少樣本]',
    'https://arxiv.org/abs/2610.07016',
    None,
)

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

# --- 2. VAND 4.0 Challenge Report — arXiv 2610.04392, benchmark/report ---

VAND4_OVERVIEW = (
    'MVTec AD 2, Kaputt 2 (VAND 4.0 Challenge benchmark report; CVPR 2026 workshop)',
    '基準/挑戰賽報告 (Benchmark / Challenge Report) — MVTec AD 2 工業缺陷分割 + Kaputt 2 零售物流',
    'VAND 4.0 Challenge Report (Detecting Defects that Matter: An Application-Driven Benchmark)',
    'MVTec, TUM, Amazon, Intel, NVIDIA (VAND 4.0 organizing committee)',
    'CVPR 2026 VAND 4.0 Workshop / arXiv 2610.04392 (2026-10-03)',
    '2026-10',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2610.04392 (Oct 3, 2026); CVPR 2026 VAND 4.0 Workshop challenge report. '
    'Industrial Track: MVTec AD 2, scored by pixel-level SegF1 (NOT AUROC). '
    'Regular setting: 9 submissions; best ~57% SegF1. Zero-shot: 5 submissions; trails by ~15 SegF1 pts. '
    'DINOv3 backbones dominate. Robustness to distribution shift flagged as open problem. '
    'Retail Track: Kaputt 2, scored by AP_any; supervised near-saturation for common defects; '
    'rare defect performance collapses (spillage ~53 AP, missing units ~27 AP). '
    'NOT a method paper — catalog reference entry. '
    'Journal harvest 2026-10-11. [基準/挑戰賽報告]',
    'https://arxiv.org/abs/2610.04392',
    'https://github.com/cvpr-vand/vand-2026',
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

# --- 3. IndusAgent — arXiv 2605.20682, open-vocabulary zero-shot ---

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
    'Headline: avg score 83.4% across datasets (metric unclear from web snippet; NOT verified from '
    'authoritative table — arXiv blocked). NON-STANDARD PROTOCOL (open-vocabulary zero-shot). '
    'Previously missed in Aug/Sep harvests. Journal harvest 2026-10-11. '
    '[基於基礎模型/VLM / 代理工具增強 / 開放詞彙 / 零樣本]',
    'https://arxiv.org/abs/2605.20682',
    None,
)

INDUSAGENT_MVTEC = (
    'MVTec AD (零樣本開放詞彙協議; NON-STANDARD PROTOCOL — open-vocabulary zero-shot)',
    '基於基礎模型 (Foundation Model-based) / VLM / 代理工具增強 (Agentic Tools) / 開放詞彙 / 零樣本',
    'IndusAgent (Reinforcing Open-Vocabulary Industrial Anomaly Detection with Agentic Tools)',
    'Rongbin Tan, Huawei Cao, Shuhan Song et al.',
    'arXiv 2605.20682 (2026-05-20)',
    '2026-05',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — open-vocabulary zero-shot protocol. arXiv HTML blocked; per-dataset metrics '
    'not accessible. Web snippet avg 83.4% (metric ambiguous). [基於基礎模型/VLM / 代理 / 零樣本]',
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

# ---------------------------------------------------------------------------
# Insert rows
# ---------------------------------------------------------------------------

def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))
    ws_all   = wb['總覽 All Papers']
    ws_mvtec = wb['MVTec AD']
    ws_visa  = wb['VisA']
    ws_ad2   = wb['MVTec AD2']
    ws_mpdd  = wb['MPDD']

    # Collect all existing paper links (col 13) across ALL sheets for global dupe check
    existing_links = set()
    for sname in wb.sheetnames:
        for row in wb[sname].iter_rows(values_only=True):
            if len(row) > 13 and row[13]:
                existing_links.add(str(row[13]).lower())

    def safe_append(ws, row_tuple, label=''):
        link = row_tuple[13]
        if link and str(link).lower() in existing_links:
            print(f'  SKIP (dupe): {row_tuple[2][:60]} → {ws.title}')
            return False
        ws.append(row_tuple)
        if link:
            existing_links.add(str(link).lower())
        tag = f' [{label}]' if label else ''
        print(f'  ADD: {row_tuple[2][:60]}{tag} → {ws.title}')
        return True

    print('=== Adding to 總覽 All Papers ===')
    safe_append(ws_all, ANCHORADAPT_OVERVIEW)
    safe_append(ws_all, VAND4_OVERVIEW)
    safe_append(ws_all, INDUSAGENT_OVERVIEW)

    print('=== Adding to MVTec AD ===')
    safe_append(ws_mvtec, ANCHORADAPT_MVTEC)
    safe_append(ws_mvtec, INDUSAGENT_MVTEC)

    print('=== Adding to VisA ===')
    safe_append(ws_visa, ANCHORADAPT_VISA)
    safe_append(ws_visa, INDUSAGENT_VISA)

    print('=== Adding to MVTec AD2 ===')
    safe_append(ws_ad2, VAND4_AD2)

    print('=== Adding to MPDD ===')
    safe_append(ws_mpdd, INDUSAGENT_MPDD)

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print('New entries: 3 papers, 8 sheet rows total (overview + per-dataset)')


if __name__ == '__main__':
    main()
