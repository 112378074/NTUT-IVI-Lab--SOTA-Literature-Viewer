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

import os, sys, json, re, time, argparse, subprocess, ssl
import urllib.request, urllib.parse
import xml.etree.ElementTree as ET

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

ARXIV_API = 'http://export.arxiv.org/api/query'
ATOM_NS   = '{http://www.w3.org/2005/Atom}'

# Search queries — keep small to stay polite to arXiv
AD_QUERIES = [
    'cat:cs.CV+AND+%28all:%22industrial+anomaly+detection%22+OR+all:%22visual+anomaly+detection%22%29',
    'cat:cs.CV+AND+all:%22MVTec%22+AND+all:%22anomaly%22',
    'cat:cs.CV+AND+all:%22VisA%22+AND+all:%22anomaly%22',
    'cat:cs.CV+AND+all:%22Real-IAD%22',
    'cat:cs.CV+AND+all:%22MVTec+LOCO%22',
]
OD_QUERIES = [
    'cat:cs.CV+AND+all:%22object+detection%22',
    'cat:cs.CV+AND+%28all:%22YOLO%22+OR+all:%22RT-DETR%22%29',
    'cat:cs.CV+AND+all:%22salient+object+detection%22',
    'cat:cs.CV+AND+%28all:%22few-shot+object+detection%22+OR+all:%22open-vocabulary+detection%22%29',
]

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

# OD category classification
OD_CATEGORIES_BY_KEYWORDS = [
    ('Real-Time OD',   ['real-time', 'real time detection', 'yolo', 'rt-detr', 'streaming detection']),
    ('RGB Salient OD', ['salient object detection', 'saliency detection', 'sod ', 'salient mask']),
    ('Few-Shot OD',    ['few-shot', 'zero-shot detection', 'open-vocabulary', 'open vocabulary',
                        'open-set', 'class-incremental', 'incremental detection', 'fsod']),
]
OD_DEFAULT_CATEGORY = 'General OD'

# OD dataset detection (subset — only the ones already in workbook)
OD_DATASETS_BY_PATTERNS = [
    ('COCO test-dev',  [r'coco\s*test-dev']),
    ('COCO 2017 val',  [r'coco\s*2017\s*val', r'coco-val', r'val2017']),
    ('LVIS v1.0 val',  [r'\blvis\b']),
    ('PASCAL VOC 2007',[r'pascal\s*voc\s*2007', r'\bvoc\s*2007']),
    ('CrowdHuman',     [r'crowdhuman']),
    ('ODinW-13',       [r'odinw-?13']),
    ('ODinW-35',       [r'odinw-?35']),
    ('DUTS-TE',        [r'duts-te', r'\bduts\b']),
    ('DUT-OMRON',      [r'dut-omron']),
    ('ECSSD',          [r'ecssd']),
    ('HKU-IS',         [r'hku-is']),
    ('PASCAL-S',       [r'pascal-s']),
    ('HRSOD',          [r'hrsod']),
    ('UHRSD',          [r'uhrsd']),
    ('DAVIS-S',        [r'davis-s']),
    ('CPPE-5',         [r'cppe-5']),
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

def fetch_arxiv(query, max_results=40, retries=3):
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
    args = parser.parse_args()

    log('=' * 60)
    log(f'Run start (mode={"dry" if args.dry else "full"})')
    log(f'Project: {PROJECT_DIR}')

    # 1. Existing IDs
    ad_ids, _ = collect_existing_arxiv_ids(AD_XLSX)
    od_ids, _ = collect_existing_arxiv_ids(OD_XLSX)
    log(f'Existing AD arxiv IDs: {len(ad_ids)} | OD arxiv IDs: {len(od_ids)}')

    # 2. Fetch + classify AD
    new_ad = []
    for q in AD_QUERIES:
        root = fetch_arxiv(q)
        for e in parse_entries(root):
            if not e['arxiv_id']:                continue
            if e['arxiv_id'] in ad_ids:          continue
            if not is_within_window(e['published'], LOOKBACK_DAYS): continue
            text = e['title'] + ' ' + e['summary']
            cat = classify_category(text, AD_CATEGORIES_BY_KEYWORDS) or AD_DEFAULT_CATEGORY
            ds  = detect_dataset(text, AD_DATASETS_BY_PATTERNS)   or AD_DEFAULT_DATASET
            method = e['title'][:90]
            ad_ids.add(e['arxiv_id'])
            new_ad.append({**e, 'category': cat, 'dataset': ds, 'method': method,
                           'note': f'Auto-fetched ({e["date"]}); 待人工核對指標'})
        time.sleep(5)
    log(f'New AD candidates: {len(new_ad)}')
    for p in new_ad:
        log(f'  AD + [{p["dataset"]}|{p["category"]}] {p["title"][:80]}')

    # 3. Fetch + classify OD
    new_od = []
    for q in OD_QUERIES:
        root = fetch_arxiv(q)
        for e in parse_entries(root):
            if not e['arxiv_id']:                continue
            if e['arxiv_id'] in od_ids:          continue
            if not is_within_window(e['published'], LOOKBACK_DAYS): continue
            text = e['title'] + ' ' + e['summary']
            cat = classify_category(text, OD_CATEGORIES_BY_KEYWORDS) or OD_DEFAULT_CATEGORY
            ds  = detect_dataset(text, OD_DATASETS_BY_PATTERNS)   or OD_DEFAULT_DATASET
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
    if not args.no_push:
        msg = f'auto: +{len(new_ad)} AD / +{len(new_od)} OD papers ({datetime.now().strftime("%Y-%m-%d")})'
        git_push(msg)

    log('Run complete')
    return 0

if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as e:
        log(f'FATAL: {type(e).__name__}: {e}')
        import traceback; log(traceback.format_exc())
        sys.exit(1)
