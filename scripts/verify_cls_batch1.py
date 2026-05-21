#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_cls_batch1.py — fill ImageNet-1K Top-1 from MambaVision Table 1 PDF."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from openpyxl import load_workbook
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
XLSX = PROJECT / 'Image_Classification_Papers_Ranking_2021_2026.xlsx'

# CLS columns: 0=subarea, 1=method, 2=authors, 3=venue, 4=date, 5=status,
# 6=Top-1, 7=Top-5, 8=Acc, 9=F1, 10=Params, 11=FLOPs, 12=notes, 13=link, 14=github

# (sheet, row, top1, params, note)
FILLS = [
    ('ImageNet-1K', 34, 83.3, '50.1M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：MambaVision-S Top-1 83.3% [Mamba-based]'),
    ('ImageNet-1K', 35, 82.3, '31.8M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：MambaVision-T Top-1 82.3% [Mamba-based]'),
    ('ImageNet-1K', 36, 83.6, '50.0M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：VMamba-S Top-1 83.6% [Mamba-based]'),
    ('ImageNet-1K', 37, 82.6, '30.0M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：VMamba-T Top-1 82.6% [Mamba-based]'),
    ('ImageNet-1K', 40, 76.1, '7.0M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：Vim-T Top-1 76.1% [Mamba-based]'),
    ('ImageNet-1K', 41, 80.5, '26.0M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：Vim-S Top-1 80.5% [Mamba-based]'),
    ('ImageNet-1K', 74, 83.9, '21.5M',
     '✅ 已驗證 (MambaVision Table 1, ImageNet-1K)：EfficientNetV2-S Top-1 83.9% [Conv-based]'),
]


def main():
    wb = load_workbook(XLSX)
    summary = []
    for sheet, row, top1, params, note in FILLS:
        s = wb[sheet]
        cur = s.cell(row=row, column=2).value
        s.cell(row=row, column=7).value = top1   # Top-1
        s.cell(row=row, column=9).value = top1   # Acc column (mirror)
        if params:
            s.cell(row=row, column=11).value = params
        s.cell(row=row, column=13).value = note
        summary.append(f'  FILL [{sheet}] r{row}: {cur} -> Top-1={top1}')
    wb.save(XLSX)
    print('Saved')
    print('\n'.join(summary))


if __name__ == '__main__':
    main()
