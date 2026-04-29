#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_papers.py
Auto-fetch new AD / OD papers from arXiv, classify, append to xlsx,
regenerate JSON + index.html, and git push.

Triggered by Windows Task Scheduler every Wed and Fri at 02:00.

Usage:
    python update_papers.py          # full run with git push
    python update_papers.py --dry    # fetch + classify + log only, no writes
    python update_papers.py --no-push # update files but skip git push
"""

import os, sys, json, re, time, argparse, subprocess, ssl, smtplib
import urllib.request, urllib.parse
import xml.etree.ElementTree as ET
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate, make_msgid

# Resolve SSL: prefer certifi's CA bundle (works in Anaconda/Windows where
# the default urllib doesn't always find a usable trust store).
try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except Exception:
    SSL_CTX = ssl.create_default_context()
from datetime import datetime, timedelta, timezone
from pathlib import Path

import openpyxl
from openpyxl import load_workbook
from copy import copy

# ====================================================================
# Configuration
# ====================================================================
PROJECT_DIR = Path(__file__).resolve().parent.parent
AD_XLSX     = PROJECT_DIR / 'AnomalyDetection_Papers_Summary_v10_20260425.xlsx'
OD_XLSX     = PROJECT_DIR / 'Object_Detection_Papers_Ranking_2021_2026.xlsx'
INDEX_HTML  = PROJECT_DIR / 'index.html'
AD_JSON     = PROJECT_DIR / 'papers_data.json'
OD_JSON     = PROJECT_DIR / 'od_data.json'
LOG_FILE    = PROJECT_DIR / 'scripts' / 'update_log.txt'
ENV_FILE    = PROJECT_DIR / 'scripts' / '.env'

ARXIV_API = 'http://export.arxiv.org/api/query'
ATOM_NS   = '{http://www.w3.org/2005/Atom}'

# Search queries — derived from CLAUDE.md (AD) and CLAUDE_updated.md (OD).
# Polite to arXiv: 5s sleep between queries, 60s timeout, 3 retries.

AD_QUERIES = [
    # Single broad AD query: anomaly + relevant dataset/method tokens
    'cat:cs.CV+AND+%28all:%22anomaly+detection%22+OR+all:%22anomaly+segmentation%22+OR+all:%22defect+detection%22+OR+all:%22defect+localization%22%29',
]

OD_QUERIES = [
    # Single broad OD query: detection + COCO/LVIS/DUTS/ODinW + key model tokens
    'cat:cs.CV+AND+%28all:%22object+detection%22+OR+all:%22salient+object%22+OR+all:%22few-shot+detection%22+OR+all:%22YOLO%22+OR+all:%22DETR%22%29',
]

# CLAUDE_updated.md §4: hard-exclude 3D / LiDAR / BEV / point-cloud / medical /
# autonomous-driving 3D / segmentation-only / tracking-only.
OD_EXCLUSION_KEYWORDS = [
    '3d object detection', '3d detection', '3d bifurcation',
    'lidar', 'point cloud', 'point-cloud', 'pointcloud',
    'bev ', "bird's-eye-view", "bird's eye view", 'birds-eye-view',
    'kitti 3d', 'nuscenes', 'waymo 3d',
    'multi-modal 3d', 'camera-only 3d',
    'medical image', 'lesion detection', 'tumor detection', 'polyp detection',
    'tooth', 'caries', 'dental', 'retinal', 'airway-tree', 'cardiac',
    'autonomous driving 3d',
    'sar object', 'remote sensing',
    'underwater', 'weed detection', 'pedestrian tracking',
    'adversarial patch', 'adversarial attack',
]
AD_EXCLUSION_KEYWORDS = [
    '3d lidar', 'lidar anomaly', 'point cloud anomaly',  # only 2D RGB / RGB-D AD
    'medical image anomaly', 'retinal abnormalit', 'lesion ',
    'video anomaly detection', 'temporal anomaly',  # not industrial visual AD
    'time series anomaly', 'log anomaly', 'network anomaly',
    'speech anomaly', 'audio anomaly',
    'surveillance', 'expressway', 'traffic anomaly',
    'fraud detection', 'cyber',
    'llm adaptation', 'language model adaptation',
]
def is_od_excluded(text):
    t = text.lower()
    return any(kw in t for kw in OD_EXCLUSION_KEYWORDS)
def is_ad_excluded(text):
    t = text.lower()
    return any(kw in t for kw in AD_EXCLUSION_KEYWORDS)

# AD category classification (priority order)
AD_CATEGORIES_BY_KEYWORDS = [
    ('基於擴散 (Diffusion)',           ['diffusion model', 'denoising diffusion', 'ddpm', 'ddim', 'score-based']),
    ('基於 NF (Normalizing Flow)',     ['normalizing flow', 'flow-based', 'fastflow', 'msflow', 'cflow', 'real-nvp']),
    ('基於資料擴增 (Data Augmentation)', ['anomaly synthesis', 'synthetic anomal', 'cutpaste', 'draem', 'glass',
                                        'pseudo anomal', 'data augmentation', 'synthesizing anomal']),
    ('基於重構 (Reconstruction)',       ['reconstruction', 'autoencoder', 'reconstruct', 'inpainting',
                                        'dinomaly', 'masked image modeling']),
    ('基於表徵 (Representation)',       ['patchcore', 'memory bank', 'student-teacher', 'student teacher',
                                        'knowledge distillation', 'efficientad', 'simplenet', 'embedding',
                                        'representation learning', 'feature distribution']),
]
AD_DEFAULT_CATEGORY = '基於表徵 (Representation)'

# AD dataset detection — order: more specific first
AD_DATASETS_BY_PATTERNS = [
    ('MVTec AD 2',  [r'mvtec\s*ad\s*2', r'mvtec\s*ad2']),
    ('MVTec LOCO',  [r'mvtec\s*loco', r'\bloco-?ad\b']),
    ('MVTec 3D',    [r'mvtec\s*3d', r'mvtec-3d-ad']),
    ('MVTec AD',    [r'mvtec\s*ad\b', r'mvtec-ad\b', r'\bmvtec\b']),
    ('VisA',        [r'\bvisa\b']),
    ('MPDD',        [r'\bmpdd\b']),
    ('BTAD',        [r'\bbtad\b', r'beantech']),
]
AD_DEFAULT_DATASET = 'MVTec AD'

# OD category classification — by alias / topic keywords. CLAUDE_updated.md §3.
OD_CATEGORIES_BY_KEYWORDS = [
    ('Few-Shot OD',    ['few-shot object detection', 'fsod', 'novel ap', 'base-to-novel',
                        'open-vocabulary detection', 'open vocabulary detection',
                        'open-vocabulary object detection', 'open-set detection',
                        'class-incremental detection', 'k-shot detection',
                        'cd-vito', 'unifs', 'grounding dino', 'odinw']),
    ('RGB Salient OD', ['salient object detection', 'saliency detection',
                        'salient region', 'birefnet', 'inspyrenet', 'duts-te',
                        'dut-omron', 'high-resolution salient']),
    ('Real-Time OD',   ['real-time object detection', 'real-time detection',
                        'yolo', 'rt-detr', 'rtdetr', 'd-fine', 'dfine ',
                        'deim ', 'rf-detr', 'rfdetr',
                        'rtmdet', 'efficientdet', 'nanodet', 'mobiledet',
                        'real time detector', 'streaming detection',
                        'edge detection', 'lightweight detection']),
]
OD_DEFAULT_CATEGORY = 'General OD'

# OD dataset detection (subset — only the ones already in workbook)
OD_DATASETS_BY_PATTERNS = [
    # Few-shot specific (most-specific first)
    ('MS-COCO 30-shot',      [r'30-shot', r'30\s+shot']),
    ('MS-COCO 10-shot',      [r'10-shot', r'10\s+shot']),
    ('MS-COCO 5-shot',       [r'5-shot',  r'5\s+shot']),
    ('MS-COCO 1-shot',       [r'1-shot',  r'1\s+shot', r'one-shot']),
    ('PASCAL VOC 2007 15+5', [r'voc\s*15\+?5', r'15\+5\s*split']),
    ('LVIS v1.0 test-dev',   [r'lvis\s*v?1\.?0?\s*test-dev', r'lvis.*test-dev']),
    ('LVIS v1.0 val',        [r'\blvis\b']),
    ('ODinW-35',             [r'odinw-?35']),
    ('ODinW-13',             [r'odinw-?13']),
    ('COCO 2017 FSOD',       [r'coco.*fsod', r'fsod.*coco']),
    # General OD
    ('COCO test-dev',        [r'coco\s*test-?dev', r'\btest-dev\b']),
    ('COCO minival',         [r'coco\s*minival', r'\bminival\b']),
    ('COCO-O',               [r'coco-o\b']),
    ('COCO 2017 val',        [r'coco\s*2017\s*val', r'coco-val', r'val2017']),
    ('COCO 2017',            [r'coco\s*2017']),
    ('PASCAL VOC 2007',      [r'pascal\s*voc\s*2007', r'\bvoc\s*2007']),
    ('CrowdHuman',           [r'crowdhuman']),
    ('GraZPEDWRI-DX',        [r'grazpedwri', r'pedwri']),
    ('CPPE-5',               [r'cppe-5', r'\bcppe\b']),
    ('Waymo 2D',             [r'waymo\s*2d']),
    # Argoverse-HD splits (most-specific first)
    ('Argoverse-HD FS Test', [r'argoverse.*full-stack.*test']),
    ('Argoverse-HD DO Test', [r'argoverse.*detection-only.*test']),
    ('Argoverse-HD FS Val',  [r'argoverse.*full-stack']),
    ('Argoverse-HD DO Val',  [r'argoverse']),
    # Salient OD
    ('DUTS-TE',              [r'duts-te', r'\bduts\b']),
    ('DUT-OMRON',            [r'dut-omron', r'omron']),
    ('ECSSD',                [r'ecssd']),
    ('HKU-IS',               [r'hku-is']),
    ('PASCAL-S',             [r'pascal-s']),
    ('HRSOD',                [r'\bhrsod\b']),
    ('UHRSD',                [r'\buhrsd\b']),
    ('DAVIS-S',              [r'davis-s']),
    ('SBU-Refine',           [r'sbu-refine']),
    ('ISTD',                 [r'\bistd\b']),
    ('CAMO-FS',              [r'camo-fs']),
]
OD_DEFAULT_DATASET = 'COCO 2017 val'

# How far back to look for new papers
LOOKBACK_DAYS = 7

# ====================================================================
# Utilities
# ====================================================================
def log(msg):
    line = f'[{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}] {msg}'
    print(line, flush=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def fetch_arxiv(query, max_results=200, retries=3):
    url = f'{ARXIV_API}?search_query={query}&sortBy=submittedDate&sortOrder=descending&max_results={max_results}'
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'NTUT-AIL-PaperBot/1.0'})
            with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as r:
                return ET.fromstring(r.read())
        except Exception as e:
            last_err = e
            wait = 5 * (attempt + 1)  # 5s, 10s, 15s
            log(f'  arxiv fetch retry {attempt+1}/{retries} after {wait}s ({type(e).__name__})')
            time.sleep(wait)
    log(f'  arxiv fetch failed for query [{query[:60]}…]: {last_err}')
    return None

def parse_entries(root):
    if root is None:
        return []
    out = []
    for e in root.findall(f'{ATOM_NS}entry'):
        eid = (e.findtext(f'{ATOM_NS}id') or '').strip()
        title = re.sub(r'\s+', ' ', e.findtext(f'{ATOM_NS}title') or '').strip()
        summary = re.sub(r'\s+', ' ', e.findtext(f'{ATOM_NS}summary') or '').strip()
        published = (e.findtext(f'{ATOM_NS}published') or '').strip()
        authors = [a.findtext(f'{ATOM_NS}name') for a in e.findall(f'{ATOM_NS}author')]
        authors = [a for a in authors if a]
        m = re.search(r'arxiv\.org/abs/(\d+\.\d+)', eid)
        arxiv_id = m.group(1) if m else None
        out.append({
            'arxiv_id': arxiv_id,
            'title': title,
            'summary': summary,
            'authors': ', '.join(authors[:8]),
            'published': published,
            'date': published[:7] if published else '',
            'link': f'https://arxiv.org/abs/{arxiv_id}' if arxiv_id else eid,
        })
    return out

def classify_category(text, mapping):
    text_l = text.lower()
    for label, kws in mapping:
        for kw in kws:
            if kw and kw in text_l:
                return label
    return None

def detect_dataset(text, mapping):
    text_l = text.lower()
    for ds, patterns in mapping:
        for p in patterns:
            if re.search(p, text_l):
                return ds
    return None

def is_within_window(published_iso, days):
    try:
        dt = datetime.strptime(published_iso[:10], '%Y-%m-%d').replace(tzinfo=timezone.utc)
    except Exception:
        return False
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    return dt >= cutoff

# ====================================================================
# Existing-papers index (to dedupe)
# ====================================================================
def collect_existing_arxiv_ids(xlsx_path):
    ids = set()
    titles = set()
    if not xlsx_path.exists():
        return ids, titles
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        for sn in wb.sheetnames:
            sh = wb[sn]
            for row in sh.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str):
                        for m in re.finditer(r'arxiv\.org/abs/(\d+\.\d+)', v):
                            ids.add(m.group(1))
                # Method names sometimes appear in column 3 (idx 2). Capture for fallback dedupe.
                if len(row) >= 3 and isinstance(row[2], str):
                    titles.add(row[2].strip().lower())
    finally:
        wb.close()
    return ids, titles

# ====================================================================
# Excel append helpers
# ====================================================================
def find_last_row(sh, key_col=1):
    last = sh.max_row
    while last > 1 and sh.cell(row=last, column=key_col).value in (None, ''):
        last -= 1
    return last

def style_from(sh, src_row, dest_row, ncols):
    for c in range(1, ncols + 1):
        s = sh.cell(row=src_row, column=c)
        d = sh.cell(row=dest_row, column=c)
        if s.has_style:
            d.font          = copy(s.font)
            d.fill          = copy(s.fill)
            d.border        = copy(s.border)
            d.alignment     = copy(s.alignment)
            d.number_format = s.number_format
            d.protection    = copy(s.protection)

def append_ad_row(wb, dataset_sheet, paper):
    """Append a row to a per-dataset AD sheet matching its 15-col format."""
    if dataset_sheet not in wb.sheetnames:
        return False
    sh = wb[dataset_sheet]
    last = find_last_row(sh, key_col=3)  # method col
    new_row = last + 1
    # Columns: 資料集 / 類別 / 方法 / 作者 / 發表 / 年月 / 狀態 / I-AUROC / P-AUROC / P-AP / P-PRO / FPS / 備註 / 連結 / GitHub
    sh.cell(row=new_row, column=1).value = dataset_sheet if dataset_sheet != 'MVTec AD2' else 'MVTec AD 2'
    sh.cell(row=new_row, column=2).value = paper['category']
    sh.cell(row=new_row, column=3).value = paper['method']
    sh.cell(row=new_row, column=4).value = paper['authors']
    sh.cell(row=new_row, column=5).value = 'arXiv preprint'
    sh.cell(row=new_row, column=6).value = paper['date']
    sh.cell(row=new_row, column=7).value = '預印本'
    # Metrics blank — to be filled in manually after review
    for c in range(8, 13):
        sh.cell(row=new_row, column=c).value = None
    sh.cell(row=new_row, column=13).value = (paper.get('note') or '自動抓取 (Auto-fetched)') + ' [' + paper['category'] + ']'
    sh.cell(row=new_row, column=14).value = paper['link']
    sh.cell(row=new_row, column=15).value = None
    style_from(sh, last if last >= 3 else 3, new_row, 15)
    return True

def append_od_dataset_row(wb, dataset_sheet, paper):
    """Append to an OD per-dataset sheet (13 cols)."""
    if dataset_sheet not in wb.sheetnames:
        return False
    sh = wb[dataset_sheet]
    last = find_last_row(sh, key_col=2)  # method col
    new_row = last + 1
    # Cols: 類別 / 方法 / 作者 / 發表 / 年月 / 狀態 / mAP / AP / FPS / params / 備註 / 連結 / GitHub
    sh.cell(row=new_row, column=1).value = paper['category']
    sh.cell(row=new_row, column=2).value = paper['method']
    sh.cell(row=new_row, column=3).value = paper['authors']
    sh.cell(row=new_row, column=4).value = 'arXiv'
    sh.cell(row=new_row, column=5).value = paper['date']
    sh.cell(row=new_row, column=6).value = 'Preprint'
    for c in range(7, 11):
        sh.cell(row=new_row, column=c).value = None
    sh.cell(row=new_row, column=11).value = paper.get('note') or 'Auto-fetched from arXiv'
    sh.cell(row=new_row, column=12).value = paper['link']
    sh.cell(row=new_row, column=13).value = None
    style_from(sh, last if last >= 2 else 2, new_row, 13)
    return True

def append_od_all_papers_row(wb, paper):
    sn = 'OD all papers'
    if sn not in wb.sheetnames:
        return False
    sh = wb[sn]
    last = find_last_row(sh, key_col=2)
    new_row = last + 1
    # Cols: 類別 / 方法 / 作者 / 發表 / 年月 / 備註 / 連結 / GitHub
    sh.cell(row=new_row, column=1).value = paper['category']
    sh.cell(row=new_row, column=2).value = paper['method']
    sh.cell(row=new_row, column=3).value = paper['authors']
    sh.cell(row=new_row, column=4).value = 'arXiv'
    sh.cell(row=new_row, column=5).value = paper['date']
    sh.cell(row=new_row, column=6).value = paper.get('note') or 'Auto-fetched from arXiv'
    sh.cell(row=new_row, column=7).value = paper['link']
    sh.cell(row=new_row, column=8).value = None
    style_from(sh, last if last >= 1 else 1, new_row, 8)
    return True

# ====================================================================
# JSON regeneration (mirrors the original extraction logic)
# ====================================================================
import math
def _clean(v):
    if v is None: return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    if isinstance(v, str) and v.strip() == '': return None
    return v

def regenerate_ad_json():
    import pandas as pd
    xl = pd.ExcelFile(AD_XLSX)
    sheets = ['MVTec AD', 'MVTec LOCO', 'MVTec AD2', 'VisA', 'MPDD', 'BTAD', 'MVTec 3D']
    rows = []
    for sn in sheets:
        if sn not in xl.sheet_names:
            continue
        df = pd.read_excel(xl, sheet_name=sn, header=2)
        df.columns = ['dataset','category','method','authors','venue','date','status',
                      'i_auroc','p_auroc','p_ap','p_pro','fps','notes','link','github']
        df = df.dropna(subset=['method'])
        for _, r in df.iterrows():
            rec = {k: _clean(v) for k, v in r.items()}
            if rec.get('method'):
                if rec.get('dataset') == 'MVTec AD2':
                    rec['dataset'] = 'MVTec AD 2'
                rows.append(rec)
    AD_JSON.write_text(json.dumps(rows, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    return rows

def regenerate_od_json():
    import pandas as pd
    xl = pd.ExcelFile(OD_XLSX)
    # Index → category map
    ds_cat = {}
    if 'Index' in xl.sheet_names:
        idx = pd.read_excel(xl, sheet_name='Index')
        for _, r in idx.iterrows():
            cat = _clean(r.get('分類 / Category'))
            sheet = _clean(r.get('資料集 Sheet'))
            if cat and sheet:
                ds_cat[sheet] = cat
    # All papers
    all_papers = []
    if 'OD all papers' in xl.sheet_names:
        ap = pd.read_excel(xl, sheet_name='OD all papers')
        for _, row in ap.iterrows():
            rec = {
                'category': _clean(row.get('類別')),
                'method':   _clean(row.get('方法')),
                'authors':  _clean(row.get('作者')),
                'venue':    _clean(row.get('發表')),
                'date':     str(_clean(row.get('年月')) or ''),
                'notes':    _clean(row.get('備註(特色/based)')),
                'link':     _clean(row.get('連結')),
                'github':   _clean(row.get('GitHub')),
            }
            if rec['method']:
                all_papers.append(rec)
    # Dataset rows
    dataset_sheets = [s for s in xl.sheet_names if s not in ('OD all papers', 'Index')]
    rows = []
    for sn in dataset_sheets:
        df = pd.read_excel(xl, sheet_name=sn)
        if '方法' not in df.columns: continue
        for _, row in df.iterrows():
            rec = {
                'dataset': sn,
                'dataset_category': ds_cat.get(sn, ''),
                'category': _clean(row.get('類別')),
                'method':   _clean(row.get('方法')),
                'authors':  _clean(row.get('作者')),
                'venue':    _clean(row.get('發表')),
                'date':     str(_clean(row.get('年月')) or ''),
                'status':   _clean(row.get('狀態')),
                'mAP':      _clean(row.get('mAP')),
                'AP':       _clean(row.get('AP')),
                'FPS':      _clean(row.get('FPS')),
                'params':   _clean(row.get('params')),
                'notes':    _clean(row.get('備註(特色/based)')),
                'link':     _clean(row.get('連結')),
                'github':   _clean(row.get('GitHub')),
            }
            if rec['method']:
                rows.append(rec)
    payload = {
        'datasets': dataset_sheets,
        'dataset_category_map': ds_cat,
        'all_papers': all_papers,
        'rows': rows,
    }
    OD_JSON.write_text(json.dumps(payload, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    return payload

def reinject_html(ad_data, od_data):
    html = INDEX_HTML.read_text(encoding='utf-8')
    ad_js = json.dumps(ad_data, ensure_ascii=False, separators=(',', ':'))
    od_js = json.dumps(od_data, ensure_ascii=False, separators=(',', ':'))
    new_html = re.sub(r'const AD_RAW = (\[[\s\S]*?\]);',
                      lambda _: 'const AD_RAW = ' + ad_js + ';', html, count=1)
    new_html = re.sub(r'const OD_DATA = (\{[\s\S]*?\});\s*\n',
                      lambda _: 'const OD_DATA = ' + od_js + ';\n', new_html, count=1)
    INDEX_HTML.write_text(new_html, encoding='utf-8')

# ====================================================================
# Email notification
# ====================================================================
def load_env(path):
    """Read simple KEY=VALUE pairs from a .env file (no quoting / interpolation)."""
    cfg = {}
    if not path.exists():
        return cfg
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        k, v = line.split('=', 1)
        cfg[k.strip()] = v.strip().strip('"').strip("'")
    return cfg

def render_email_body(new_ad, new_od, ran_at, pushed):
    """Plain-text body listing the new papers."""
    lines = []
    lines.append(f'NTUT 自動化檢測實驗室 — 文獻自動更新通知')
    lines.append(f'執行時間: {ran_at}')
    lines.append(f'本次新增: AD {len(new_ad)} 篇 / OD {len(new_od)} 篇')
    lines.append(f'GitHub 推送: {"成功" if pushed else "未推送"}')
    lines.append('')
    lines.append('=' * 64)
    lines.append('Anomaly Detection 新增論文')
    lines.append('=' * 64)
    if new_ad:
        for i, p in enumerate(new_ad, 1):
            lines.append(f'{i:>2}. [{p["dataset"]} | {p["category"]}]')
            lines.append(f'    {p["title"]}')
            if p.get('authors'):
                lines.append(f'    Authors : {p["authors"]}')
            lines.append(f'    arXiv   : {p["link"]}')
            lines.append(f'    Date    : {p["date"]}')
            lines.append('')
    else:
        lines.append('  (本次無新增)')
        lines.append('')

    lines.append('=' * 64)
    lines.append('Object Detection 新增論文')
    lines.append('=' * 64)
    if new_od:
        for i, p in enumerate(new_od, 1):
            lines.append(f'{i:>2}. [{p["dataset"]} | {p["category"]}]')
            lines.append(f'    {p["title"]}')
            if p.get('authors'):
                lines.append(f'    Authors : {p["authors"]}')
            lines.append(f'    arXiv   : {p["link"]}')
            lines.append(f'    Date    : {p["date"]}')
            lines.append('')
    else:
        lines.append('  (本次無新增)')
        lines.append('')

    lines.append('-' * 64)
    lines.append('提示: 自動分類為 best-effort 啟發式, 指標數字 (AUROC / mAP 等)')
    lines.append('需手動填入。請參閱 scripts/update_log.txt 確認分類正確。')
    lines.append('網站: https://112378074.github.io/NTUT-IVI-Lab--SOTA-Literature-Viewer/')
    return '\n'.join(lines)

def render_email_html(new_ad, new_od, ran_at, pushed):
    def section(title, papers, cat_color):
        if not papers:
            return f'<h3 style="color:#0f172a">{title}</h3><p style="color:#64748b">本次無新增</p>'
        rows = ''
        for i, p in enumerate(papers, 1):
            rows += (f'<tr style="border-top:1px solid #e2e8f0">'
                     f'<td style="padding:8px;color:#64748b;width:32px;text-align:right">{i}</td>'
                     f'<td style="padding:8px">'
                     f'<div style="font-weight:600;color:#0f172a">{p["title"]}</div>'
                     f'<div style="color:#475569;font-size:13px;margin-top:2px">'
                     f'<span style="background:{cat_color};color:#fff;padding:1px 7px;border-radius:10px;font-size:11px;margin-right:6px">{p["category"]}</span>'
                     f'<span style="background:#e2e8f0;color:#334155;padding:1px 7px;border-radius:10px;font-size:11px;margin-right:6px">{p["dataset"]}</span>'
                     f'<span style="color:#64748b">{p["date"]}</span></div>'
                     f'<div style="color:#64748b;font-size:12px;margin-top:4px">'
                     f'{(p.get("authors") or "")[:120]}</div>'
                     f'<div style="margin-top:4px"><a href="{p["link"]}" style="color:#2563eb;font-size:12px">arXiv ↗</a></div>'
                     f'</td></tr>')
        return (f'<h3 style="color:#0f172a;margin-bottom:8px">{title}</h3>'
                f'<table style="width:100%;border-collapse:collapse;background:#fff;'
                f'border:1px solid #e2e8f0;border-radius:8px;overflow:hidden">{rows}</table>')

    push_pill = ('<span style="background:#dcfce7;color:#166534;padding:2px 8px;border-radius:10px;font-size:12px">已推送 GitHub</span>'
                 if pushed
                 else '<span style="background:#fef3c7;color:#854d0e;padding:2px 8px;border-radius:10px;font-size:12px">未推送</span>')
    html = f'''<!doctype html><html><body style="font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#f6f8fc;padding:20px;color:#1a2233">
<div style="max-width:780px;margin:0 auto">
  <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:16px">
    <div style="font-size:12px;color:#2563eb;font-weight:600;margin-bottom:8px">NTUT · IIM · Automated Inspection Lab</div>
    <h2 style="margin:0 0 6px;color:#0f172a">文獻自動更新通知</h2>
    <div style="color:#64748b;font-size:13px">執行時間 {ran_at} · 新增 AD {len(new_ad)} 篇 / OD {len(new_od)} 篇 · {push_pill}</div>
  </div>
  <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:16px">
    {section("Anomaly Detection 新增論文", new_ad, "#2563eb")}
  </div>
  <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px;margin-bottom:16px">
    {section("Object Detection 新增論文", new_od, "#7c3aed")}
  </div>
  <div style="color:#64748b;font-size:12px;text-align:center;line-height:1.6;padding:8px 0 16px">
    自動分類為 best-effort 啟發式; 指標數字需手動填入。<br>
    <a href="https://112378074.github.io/NTUT-IVI-Lab--SOTA-Literature-Viewer/" style="color:#2563eb">前往網站</a>
  </div>
</div></body></html>'''
    return html

def send_notification(new_ad, new_od, ran_at, pushed):
    cfg = load_env(ENV_FILE)
    host = cfg.get('SMTP_HOST', 'smtp.gmail.com')
    port = int(cfg.get('SMTP_PORT', '587'))
    user = cfg.get('SMTP_USER', '')
    pwd  = cfg.get('SMTP_PASSWORD', '')
    to   = cfg.get('NOTIFY_TO', 'azaz31855@gmail.com')

    if not user or not pwd:
        log(f'  email skipped: SMTP_USER / SMTP_PASSWORD not set in {ENV_FILE}')
        return False

    subject = f'[AIL Auto-Update] +{len(new_ad)} AD / +{len(new_od)} OD ({datetime.now().strftime("%Y-%m-%d")})'
    text = render_email_body(new_ad, new_od, ran_at, pushed)
    html = render_email_html(new_ad, new_od, ran_at, pushed)

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From']    = user
    msg['To']      = to
    msg['Date']    = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid()
    msg.attach(MIMEText(text, 'plain', 'utf-8'))
    msg.attach(MIMEText(html, 'html',  'utf-8'))

    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            s.starttls(context=SSL_CTX)
            s.ehlo()
            s.login(user, pwd)
            s.sendmail(user, [to], msg.as_string())
        log(f'  email sent to {to}')
        return True
    except Exception as e:
        log(f'  email send failed: {type(e).__name__}: {e}')
        return False

# ====================================================================
# Git push
# ====================================================================
def git_push(message):
    try:
        cwd = str(PROJECT_DIR)
        files = [INDEX_HTML.name, AD_XLSX.name, OD_XLSX.name, AD_JSON.name, OD_JSON.name]
        subprocess.run(['git', 'add'] + files, cwd=cwd, check=True)
        # Check if anything to commit
        diff = subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=cwd)
        if diff.returncode == 0:
            log('  no changes to commit')
            return False
        subprocess.run(['git', 'commit', '-m', message], cwd=cwd, check=True)
        subprocess.run(['git', 'push'], cwd=cwd, check=True)
        log('  git push completed')
        return True
    except subprocess.CalledProcessError as e:
        log(f'  git operation failed: {e}')
        return False

# ====================================================================
# Main
# ====================================================================
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry', action='store_true', help='Fetch and classify only; no writes')
    parser.add_argument('--no-push', action='store_true', help='Skip git push')
    parser.add_argument('--no-email', action='store_true', help='Skip email notification')
    args = parser.parse_args()
    ran_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    log('=' * 60)
    log(f'Run start (mode={"dry" if args.dry else "full"})')
    log(f'Project: {PROJECT_DIR}')

    # 1. Existing IDs
    ad_ids, _ = collect_existing_arxiv_ids(AD_XLSX)
    od_ids, _ = collect_existing_arxiv_ids(OD_XLSX)
    log(f'Existing AD arxiv IDs: {len(ad_ids)} | OD arxiv IDs: {len(od_ids)}')

    # 2. Fetch + classify AD — require target dataset to be mentioned (avoid noise)
    new_ad = []
    for q in AD_QUERIES:
        root = fetch_arxiv(q)
        for e in parse_entries(root):
            if not e['arxiv_id']:                continue
            if e['arxiv_id'] in ad_ids:          continue
            if not is_within_window(e['published'], LOOKBACK_DAYS): continue
            text = e['title'] + ' ' + e['summary']
            if is_ad_excluded(text):
                log(f'  AD - skipped (excluded topic): {e["title"][:80]}')
                continue
            ds = detect_dataset(text, AD_DATASETS_BY_PATTERNS)
            if not ds:
                # No target dataset mentioned -> probably off-topic
                log(f'  AD - skipped (no target dataset): {e["title"][:80]}')
                continue
            cat = classify_category(text, AD_CATEGORIES_BY_KEYWORDS) or AD_DEFAULT_CATEGORY
            method = e['title'][:90]
            ad_ids.add(e['arxiv_id'])
            new_ad.append({**e, 'category': cat, 'dataset': ds, 'method': method,
                           'note': f'Auto-fetched ({e["date"]}); 待人工核對指標'})
        time.sleep(5)
    log(f'New AD candidates: {len(new_ad)}')
    for p in new_ad:
        log(f'  AD + [{p["dataset"]}|{p["category"]}] {p["title"][:80]}')

    # 3. Fetch + classify OD — require target dataset; exclude 3D/medical/etc.
    new_od = []
    for q in OD_QUERIES:
        root = fetch_arxiv(q)
        for e in parse_entries(root):
            if not e['arxiv_id']:                continue
            if e['arxiv_id'] in od_ids:          continue
            if not is_within_window(e['published'], LOOKBACK_DAYS): continue
            text = e['title'] + ' ' + e['summary']
            if is_od_excluded(text):
                log(f'  OD - skipped (excluded topic): {e["title"][:80]}')
                continue
            ds = detect_dataset(text, OD_DATASETS_BY_PATTERNS)
            if not ds:
                log(f'  OD - skipped (no target dataset): {e["title"][:80]}')
                continue
            cat = classify_category(text, OD_CATEGORIES_BY_KEYWORDS) or OD_DEFAULT_CATEGORY
            method = e['title'][:90]
            od_ids.add(e['arxiv_id'])
            new_od.append({**e, 'category': cat, 'dataset': ds, 'method': method,
                           'note': 'Auto-fetched from arXiv; metrics To verify'})
        time.sleep(5)
    log(f'New OD candidates: {len(new_od)}')
    for p in new_od:
        log(f'  OD + [{p["dataset"]}|{p["category"]}] {p["title"][:80]}')

    if args.dry:
        log('Dry run complete; no files written')
        return 0

    if not new_ad and not new_od:
        log('No new papers found — nothing to update')
        if not args.no_email:
            send_notification([], [], ran_at, pushed=False)
        return 0

    # 4. Append AD rows
    if new_ad:
        wb = load_workbook(AD_XLSX)
        for p in new_ad:
            sheet = p['dataset']
            if sheet == 'MVTec AD 2': sheet = 'MVTec AD2'
            ok = append_ad_row(wb, sheet, p)
            if not ok:
                log(f'  AD: skipped (sheet {sheet} not found): {p["title"][:60]}')
        wb.save(AD_XLSX)
        log(f'AD xlsx updated: +{len(new_ad)} rows')

    # 5. Append OD rows
    if new_od:
        wb = load_workbook(OD_XLSX)
        for p in new_od:
            append_od_dataset_row(wb, p['dataset'], p)
            append_od_all_papers_row(wb, p)
        wb.save(OD_XLSX)
        log(f'OD xlsx updated: +{len(new_od)} rows')

    # 6. Regenerate JSON & inject HTML
    ad_data = regenerate_ad_json()
    od_data = regenerate_od_json()
    reinject_html(ad_data, od_data)
    log(f'Regenerated: {len(ad_data)} AD records, {len(od_data["all_papers"])} OD methods, {len(od_data["rows"])} OD rows')

    # 7. Git push
    pushed = False
    if not args.no_push:
        msg = f'auto: +{len(new_ad)} AD / +{len(new_od)} OD papers ({datetime.now().strftime("%Y-%m-%d")})'
        pushed = git_push(msg)

    # 8. Email notification
    if not args.no_email:
        send_notification(new_ad, new_od, ran_at, pushed)

    log('Run complete')
    return 0

if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        log(f'FATAL: {type(e).__name__}: {e}')
        import traceback; log(traceback.format_exc())
        sys.exit(1)
