#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Weekly IVAD Journal Harvest — 2026-11-22
Coverage window: 2026-11-16 to 2026-11-22

Target journals searched:
  IEEE TPAMI, IEEE TII, IEEE TNNLS, IEEE TIP, IEEE TCSVT,
  Pattern Recognition, IJCV, Computers in Industry,
  Information Fusion, Knowledge-Based Systems,
  Information Sciences, Neural Networks

Network access this run:
  - arXiv.org HTML:   BLOCKED by environment proxy (consistent with prior weeks)
  - IEEE Xplore:      BLOCKED by environment proxy
  - ScienceDirect:    BLOCKED by environment proxy
  - GitHub:           ACCESSIBLE (raw.githubusercontent.com — metric verification)
  - Web search:       ACCESSIBLE (Google snippets via WebSearch tool)

Coverage note (Nov 16–22, 2026):
  arXiv IDs for the Nov 16–22 window are approximately 2611.24xxx–2611.28xxx.
  Two papers in that range found and auto-added 2026-11-20 by the automated run:
    FBR  (2611.24541): already in MVTec 3D sheet — venue and Total Papers row missing
    AggCLIP (2611.24643): already in MVTec 3D sheet — Total Papers row missing
  One paper from the prior week (2611.13xxx, ~Nov 11) was auto-fetched into the
  WRONG sheet (MVTec AD instead of MVTec LOCO):
    CASCADE (2611.13820): corrected here

Corrections this run:
  1. CASCADE (arXiv 2611.13820, ICDM 2026):
     - Fix row in MVTec AD (wrong sheet — cascade is a logical-AD/MVTec LOCO paper)
     - Add correct entry to MVTec LOCO sheet
     - Add entry to 總覽 All Papers

  2. FBR (arXiv 2611.24541, Information Fusion 2026):
     - Venue confirmed from GitHub README (LiZhi-CASIA/FBR): "Information Fusion 2026,
       DOI 10.1016/j.inffus.2026.104644". GitHub README has no results table.
       arXiv HTML blocked. Status: ⏳ Pending for metrics.
     - Update venue in MVTec 3D row ('arXiv preprint' → 'Information Fusion 2026 …')
     - Add entry to 總覽 All Papers

  3. AggCLIP (arXiv 2611.24643):
     - Venue not confirmed (no GitHub found, no venue from search results).
     - Add entry to 總覽 All Papers (keeps '預印本' status).

  4. DART (arXiv 2611.08271):
     - Venue update: NeurIPS 2026 confirmed from neurips.cc/virtual/2026/poster/152532.
     - No code available yet ("will be released upon publication").

Pending backlog retried this run (2026-11-22):
  - FATE (2611.08140):        No GitHub found; arXiv HTML blocked → still ⏳ Pending
  - TIRAD (2611.12023):       No GitHub found; arXiv HTML blocked → still ⏳ Pending
  - PRISM (2611.08136):       No GitHub found; arXiv HTML blocked → still ⏳ Pending
  - DART (2611.08271):        NeurIPS 2026 venue confirmed; no code yet → still ⏳ Pending (metrics)
  - FoundAD (2510.01934):     GitHub README has no results table → still ⏳ Pending
  - Robust-FSAD (2510.13643): GitHub README has no results table → still ⏳ Pending

New entries to add this week: 3 entries (CASCADE → MVTec LOCO + Total;
FBR → Total; AggCLIP → Total); plus corrections to existing rows.
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
# 1. CASCADE (arXiv 2611.13820, ICDM 2026)
#    "CASCADE: Constraint-Aware SCAffolded DEcomposition for Logical Anomaly Detection"
#    Authors: Jinchan Kim, Jinhee Im, Hyunsouk Cho
#    Dataset: MVTec LOCO AD (also VID-AD) — NOT standard MVTec AD
#    Protocol: Training-free, 3-shot reference-based (NON-STANDARD)
#    Metric: AUROC (MVTec LOCO); CASCADE++ = 93.2 ± 0.8 (3 reference images)
#    Verification: arXiv HTML blocked; no GitHub found. Metrics from search
#                  snippets of arXiv PDF — NOT from an accessible authoritative
#                  table. NOT ranked per workflow rules.
# ===========================================================================

CASCADE_LOCO = (
    'MVTec LOCO AD (training-free 3-shot reference-based; NON-STANDARD protocol)',
    '邏輯異常偵測 (Logical Anomaly Detection-based) / 訓練免費 (Training-Free) '
    '/ 約束感知 (Constraint-Aware) / 分解推理 (Scaffolded Decomposition)',
    'CASCADE (Constraint-Aware SCAffolded DEcomposition for Logical Anomaly Detection)',
    'Jinchan Kim, Jinhee Im, Hyunsouk Cho',
    'ICDM 2026 / arXiv 2611.13820 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — ICDM 2026 / arXiv 2611.13820. No GitHub found; arXiv HTML '
    'blocked. Method: Training-free, 3-shot reference-based logical AD. Sorts '
    'constraints into binding vs calibration types; uses SAM3 for object grounding. '
    'CASCADE++ reports 93.2 AUROC (± 0.8 std) with 3 reference images on MVTec LOCO '
    '— from arXiv PDF search snippet, NOT from an accessible authoritative table, '
    'NOT ranked. 71.8 s/image inference. Protocol: training-free reference-based '
    'logical AD — NOT comparable to standard supervised AD on MVTec LOCO. '
    '[Logical AD / Training-Free / Constraint-Aware / ICDM 2026]',
    'https://arxiv.org/abs/2611.13820',
    None,
)

CASCADE_TOTAL = (
    'MVTec LOCO AD (training-free 3-shot reference-based; NON-STANDARD protocol)',
    '邏輯異常偵測 (Logical Anomaly Detection-based) / 訓練免費 (Training-Free) '
    '/ 約束感知 (Constraint-Aware) / 分解推理 (Scaffolded Decomposition)',
    'CASCADE (Constraint-Aware SCAffolded DEcomposition for Logical Anomaly Detection)',
    'Jinchan Kim, Jinhee Im, Hyunsouk Cho',
    'ICDM 2026 / arXiv 2611.13820 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — ICDM 2026 / arXiv 2611.13820. No GitHub found; arXiv HTML '
    'blocked. Training-free logical AD; 3-shot reference images on MVTec LOCO. '
    'NOT ranked. Protocol: training-free reference-based — NON-STANDARD. '
    'Journal harvest 2026-11-22. [Logical AD / Training-Free / ICDM 2026]',
    'https://arxiv.org/abs/2611.13820',
    None,
)

# ===========================================================================
# 2. FBR (arXiv 2611.24541, Information Fusion 2026)
#    "Fuse Before Reconstruct: A Novel Reconstruction Paradigm for Multimodal
#     Industrial Anomaly Detection"
#    Authors: Zhi Li, Hanyu Wang, Qiyu Chen, Fei Shen, Zhengtao Zhang
#    Dataset: MVTec 3D-AD, Eyecandies
#    Venue confirmed from GitHub README (LiZhi-CASIA/FBR):
#      "Information Fusion 2026, DOI 10.1016/j.inffus.2026.104644"
#    GitHub README has no numeric results table (only training/eval commands).
#    arXiv HTML blocked. Metrics: ⏳ Pending.
# ===========================================================================

FBR_TOTAL = (
    'MVTec 3D-AD, Eyecandies (3D/RGB-D multimodal AD)',
    '基於重構 (Reconstruction-based) / 3D/RGB-D融合 (3D/RGB-D Fusion-based) '
    '/ 多模態先融合後重構 (Fuse-Then-Reconstruct)',
    'FBR (Fuse Before Reconstruct: A Novel Reconstruction Paradigm for Multimodal '
    'Industrial Anomaly Detection)',
    'Zhi Li, Hanyu Wang, Qiyu Chen, Fei Shen, Zhengtao Zhang',
    'Information Fusion 2026 (DOI: 10.1016/j.inffus.2026.104644) / arXiv 2611.24541',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — Information Fusion 2026, DOI 10.1016/j.inffus.2026.104644 '
    '(venue confirmed from GitHub README LiZhi-CASIA/FBR). GitHub README has '
    'no numeric results table (setup/training commands only). arXiv HTML '
    'blocked. Method: RGB+3D (point cloud) multimodal fusion; fuses cross-modal '
    'features before reconstruction. Backbone: DINO ViT-B/8 (RGB) + Point-MAE '
    'Point Transformer (3D). Datasets: MVTec 3D-AD, Eyecandies. '
    'NOT ranked. Journal harvest 2026-11-22. '
    '[Reconstruction / 3D-RGB-D Fusion / Information Fusion 2026]',
    'https://arxiv.org/abs/2611.24541',
    'https://github.com/LiZhi-CASIA/FBR',
)

# ===========================================================================
# 3. AggCLIP (arXiv 2611.24643)
#    "Aggregating CLIP for Zero-Shot Multimodal Anomaly Detection"
#    Authors: Zhi Li, Hanyu Wang, Fei Shen, Zhengtao Zhang (same group as FBR)
#    Dataset: MVTec 3D-AD, Eyecandies (zero-shot multimodal)
#    Venue: not confirmed (no GitHub found, no venue confirmed by search)
#    Status: ⏳ Pending
# ===========================================================================

AGGCLIP_TOTAL = (
    'MVTec 3D-AD, Eyecandies (zero-shot multimodal AD)',
    '視覺語言模型 (VLM-based) / CLIP-based / 零樣本 (Zero-Shot AD) / 多模態 (Multimodal)',
    'AggCLIP (Aggregating CLIP for Zero-Shot Multimodal Anomaly Detection)',
    'Zhi Li, Hanyu Wang, Fei Shen, Zhengtao Zhang',
    'arXiv 2611.24643 (2026-11)',
    '2026-11',
    '⏳ Pending',
    None, None, None, None, None,
    '⏳ Pending — arXiv 2611.24643 (Nov 2026). No GitHub found (tried '
    'LiZhi-CASIA/AggCLIP); venue not confirmed. Method: Aggregates CLIP features '
    'for zero-shot multimodal AD on MVTec 3D-AD and Eyecandies. Reports +1.4% '
    'O-AUROC over PointAD on MVTec 3D-AD (from search snippet — NOT from an '
    'accessible authoritative table, NOT ranked). '
    'Protocol: zero-shot — NOT comparable to standard supervised AD. '
    'Journal harvest 2026-11-22. [VLM / CLIP / Zero-shot / Multimodal 3D]',
    'https://arxiv.org/abs/2611.24643',
    None,
)


# ---------------------------------------------------------------------------
# Update helpers
# ---------------------------------------------------------------------------

def find_row_by_link(ws, link):
    """Find row index (1-based) by arXiv link in col 14."""
    link_lower = link.lower()
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if len(row) > 13 and row[13] and str(row[13]).lower() == link_lower:
            return i
    return None


def sheet_has_link(ws, link):
    return find_row_by_link(ws, link) is not None


def safe_append(ws, row_tuple):
    link = row_tuple[13] or ''
    method = str(row_tuple[2])
    if link and sheet_has_link(ws, link):
        # Also check method name to distinguish per-dataset rows
        for row in ws.iter_rows(values_only=True):
            if len(row) > 13 and row[13] and str(row[13]).lower() == link.lower():
                # Check dataset col too
                if str(row[0]).lower() == str(row_tuple[0]).lower():
                    print(f'  SKIP (dupe): {method[:55]} → {ws.title}')
                    return False
    ws.append(row_tuple)
    print(f'  ADD: {method[:55]} → {ws.title}')
    return True


def update_cell(ws, row_idx, col_idx, new_value):
    """Update a single cell (1-based row/col indices)."""
    ws.cell(row=row_idx, column=col_idx, value=new_value)


def main():
    wb = openpyxl.load_workbook(str(AD_XLSX))

    ws_all   = None
    ws_loco  = wb['MVTec LOCO']
    ws_mvtec = wb['MVTec AD']
    ws_3d    = wb['MVTec 3D']

    for name in wb.sheetnames:
        if 'All' in name or '總' in name:
            ws_all = wb[name]
            break
    if ws_all is None:
        ws_all = wb.active

    added = 0
    updated = 0

    # -----------------------------------------------------------------------
    # 1. Fix CASCADE in MVTec AD sheet — it was auto-fetched to the wrong sheet
    #    (CASCADE is a MVTec LOCO paper, not standard MVTec AD)
    # -----------------------------------------------------------------------
    print('=== Fixing CASCADE row in MVTec AD (wrong placement) ===')
    cascade_link = 'https://arxiv.org/abs/2611.13820'
    cascade_row_idx = find_row_by_link(ws_mvtec, cascade_link)
    if cascade_row_idx:
        # Col 1 = 資料集, Col 5 = 發表, Col 7 = 狀態, Col 13 = 備註
        update_cell(ws_mvtec, cascade_row_idx, 1,
                    'MVTec LOCO AD (誤植至本頁 — 已於 MVTec LOCO 頁補入正確條目)')
        update_cell(ws_mvtec, cascade_row_idx, 2,
                    '邏輯異常偵測 (Logical Anomaly Detection-based) / 訓練免費 '
                    '(Training-Free) / ICDM 2026')
        update_cell(ws_mvtec, cascade_row_idx, 5,
                    'ICDM 2026 / arXiv 2611.13820 (2026-11)')
        update_cell(ws_mvtec, cascade_row_idx, 7, '⏳ Pending')
        update_cell(ws_mvtec, cascade_row_idx, 13,
                    '⚠️ 此條目誤植至 MVTec AD 頁 — CASCADE 為 MVTec LOCO '
                    'logical AD 論文，非 standard MVTec AD。正確條目已補入 MVTec LOCO 頁。'
                    'ICDM 2026 / arXiv 2611.13820. No GitHub found; arXiv HTML blocked. '
                    'Protocol: training-free 3-shot reference-based — NON-STANDARD. '
                    'NOT ranked. Corrected 2026-11-22. [Logical AD / ICDM 2026]')
        print(f'  UPDATED: CASCADE row {cascade_row_idx} in MVTec AD (wrong-placement note)')
        updated += 1
    else:
        print('  CASCADE not found in MVTec AD (already corrected or missing)')

    # -----------------------------------------------------------------------
    # 2. Add CASCADE to MVTec LOCO and Total Papers
    # -----------------------------------------------------------------------
    print('=== Adding CASCADE (arXiv 2611.13820, ICDM 2026) to correct sheets ===')
    added += safe_append(ws_loco, CASCADE_LOCO)
    added += safe_append(ws_all,  CASCADE_TOTAL)

    # -----------------------------------------------------------------------
    # 3. Update FBR venue in MVTec 3D (arXiv preprint → Information Fusion 2026)
    # -----------------------------------------------------------------------
    print('=== Updating FBR venue in MVTec 3D ===')
    fbr_link = 'https://arxiv.org/abs/2611.24541'
    fbr_row_idx = find_row_by_link(ws_3d, fbr_link)
    if fbr_row_idx:
        update_cell(ws_3d, fbr_row_idx, 5,
                    'Information Fusion 2026 (DOI: 10.1016/j.inffus.2026.104644) '
                    '/ arXiv 2611.24541')
        update_cell(ws_3d, fbr_row_idx, 7, '⏳ Pending')
        update_cell(ws_3d, fbr_row_idx, 13,
                    '⏳ Pending — Information Fusion 2026, '
                    'DOI 10.1016/j.inffus.2026.104644 (venue confirmed from '
                    'GitHub README LiZhi-CASIA/FBR). GitHub README has no numeric '
                    'results table (setup/training commands only). arXiv HTML blocked. '
                    'Method: RGB+3D multimodal fusion before reconstruction. '
                    'Backbone: DINO ViT-B/8 + Point-MAE Point Transformer. '
                    'Datasets: MVTec 3D-AD, Eyecandies. NOT ranked. '
                    'Updated 2026-11-22. [Reconstruction / 3D-RGB-D Fusion / IF 2026]')
        update_cell(ws_3d, fbr_row_idx, 15, 'https://github.com/LiZhi-CASIA/FBR')
        print(f'  UPDATED: FBR row {fbr_row_idx} in MVTec 3D (venue: IF 2026)')
        updated += 1
    else:
        print('  FBR not found in MVTec 3D (may need manual check)')

    # -----------------------------------------------------------------------
    # 4. Add FBR to Total Papers (missing from 總覽)
    # -----------------------------------------------------------------------
    print('=== Adding FBR to 總覽 All Papers ===')
    added += safe_append(ws_all, FBR_TOTAL)

    # -----------------------------------------------------------------------
    # 5. Add AggCLIP to Total Papers (missing from 總覽)
    # -----------------------------------------------------------------------
    print('=== Adding AggCLIP to 總覽 All Papers ===')
    added += safe_append(ws_all, AGGCLIP_TOTAL)

    # -----------------------------------------------------------------------
    # 6. Update DART venue in Total Papers (arXiv 2611.08271 → NeurIPS 2026)
    # -----------------------------------------------------------------------
    print('=== Updating DART venue in 總覽 All Papers ===')
    dart_link = 'https://arxiv.org/abs/2611.08271'
    dart_row_idx = find_row_by_link(ws_all, dart_link)
    if dart_row_idx:
        # Col 5 = 發表
        old_venue = ws_all.cell(row=dart_row_idx, column=5).value
        update_cell(ws_all, dart_row_idx, 5,
                    'NeurIPS 2026 / arXiv 2611.08271 (2026-11)')
        old_note = ws_all.cell(row=dart_row_idx, column=13).value or ''
        if 'NeurIPS' not in old_note:
            new_note = (
                '⏳ Pending — NeurIPS 2026 (confirmed: neurips.cc/virtual/2026/'
                'poster/152532) / arXiv 2611.08271. No GitHub code yet ("code will '
                'be available upon publication" per NeurIPS poster page). arXiv HTML '
                'blocked. Method: plug-in GAD framework; learns to separate '
                'anomaly-relevant from domain-specific residuals. Evaluated on '
                'MVTec AD, VisA, BraTS. Protocol: Generalist AD / cross-domain — '
                'NOT comparable to standard single-dataset unsupervised AD results. '
                'NOT ranked. Updated 2026-11-22. [Representation / Residual Transfer '
                '/ Generalist AD / NeurIPS 2026]'
            )
            update_cell(ws_all, dart_row_idx, 13, new_note)
        print(f'  UPDATED: DART row {dart_row_idx} venue → NeurIPS 2026 '
              f'(was: {old_venue})')
        updated += 1
    else:
        print('  DART not found in 總覽 All Papers')

    wb.save(str(AD_XLSX))
    print(f'\n✅ Saved: {AD_XLSX}')
    print(f'New entries added: {added} sheet rows')
    print(f'Existing rows updated: {updated}')
    print()
    print('Summary of this harvest (2026-11-22):')
    print()
    print('=== NEW ENTRIES ===')
    print('  CASCADE (2611.13820, ICDM 2026):')
    print('    → MVTec LOCO sheet: ⏳ Pending (no GitHub; arXiv HTML blocked)')
    print('    → 總覽 All Papers:   ⏳ Pending')
    print('    → Fixed wrong entry in MVTec AD (had been auto-fetched to wrong sheet)')
    print()
    print('  FBR (2611.24541, Information Fusion 2026):')
    print('    → 總覽 All Papers:   ⏳ Pending (GitHub README has no metrics table)')
    print('    → MVTec 3D:          Venue updated: "arXiv preprint" → "IF 2026"')
    print()
    print('  AggCLIP (2611.24643):')
    print('    → 總覽 All Papers:   ⏳ Pending (venue not confirmed; no GitHub)')
    print()
    print('=== PENDING BACKLOG (no new verification sources found) ===')
    print('  DART (2611.08271):        NeurIPS 2026 venue confirmed; no code yet')
    print('  FATE (2611.08140):        ⏳ Pending — no GitHub; arXiv HTML blocked')
    print('  TIRAD (2611.12023):       ⏳ Pending — no GitHub; arXiv HTML blocked')
    print('  PRISM (2611.08136):       ⏳ Pending — no GitHub; arXiv HTML blocked')
    print('  FoundAD (2510.01934):     ⏳ Pending — GitHub README no metrics table')
    print('  Robust-FSAD (2510.13643): ⏳ Pending — GitHub README no metrics table')


if __name__ == '__main__':
    main()
