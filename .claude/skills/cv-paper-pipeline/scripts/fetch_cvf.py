#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_cvf.py — CVF Open Access scraper for the cv-paper-pipeline skill.

Fetches papers from CVF Open Access (CVPR / ICCV / ECCV / WACV proceedings),
filters them by domain keyword (AD / OD / Classification), classifies each
paper, and writes a CSV that can be reviewed and pasted into the Excel
workbooks.

CVF Open Access is fully public HTML — no login required.

Usage:
    python fetch_cvf.py --conf CVPR --year 2025
    python fetch_cvf.py --conf CVPR --year 2025 --conf ICCV --year 2025
    python fetch_cvf.py --conf CVPR --year 2025 --domain ad
    python fetch_cvf.py --conf CVPR --year 2025 --out cvpr2025.csv

Output CSV columns:
    domain, conference, year, title, authors, paper_url, dataset_guess,
    category_guess, abstract_snippet
"""

import argparse, csv, re, ssl, sys, time, urllib.request
from pathlib import Path

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except Exception:
    SSL_CTX = ssl.create_default_context()

CVF_BASE = 'https://openaccess.thecvf.com'
UA = 'NTUT-AIL-PaperBot/1.0 (research literature review; contact azaz31855@gmail.com)'

# ---- Domain keyword classifiers --------------------------------------------
AD_KW = ['anomaly detection', 'anomaly segmentation', 'defect detection',
         'defect localization', 'mvtec', 'visa', 'mpdd', 'btad',
         'industrial anomaly', 'unsupervised anomaly', 'surface defect']
OD_KW = ['object detection', 'salient object', 'few-shot detection',
         'real-time detection', '3d object detection', 'detr', 'yolo',
         'open-vocabulary detection']
CLS_KW = ['image classification', 'fine-grained classification',
          'few-shot classification', 'semi-supervised classification',
          'long-tailed classification', 'small data classification',
          'image recognition']

# Topics to drop (medical-only, segmentation-only, tracking-only, etc.)
EXCLUDE_KW = ['medical image segmentation', 'cell segmentation',
              'speech', 'audio-only', 'point cloud registration']


def fetch(url, retries=3):
    last = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': UA})
            with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as r:
                return r.read().decode('utf-8', errors='replace')
        except Exception as e:
            last = e
            time.sleep(5 * (attempt + 1))
    print(f'  fetch failed: {url} ({last})', file=sys.stderr)
    return None


def parse_conference(conf, year):
    """Return list of dicts: {title, authors, paper_url} for one conference."""
    # CVF lists proceedings either at /{CONF}{YEAR}?day=all or /{CONF}{YEAR}
    out = []
    for suffix in (f'{conf}{year}?day=all', f'{conf}{year}'):
        html = fetch(f'{CVF_BASE}/{suffix}')
        if not html:
            continue
        # <dt class="ptitle"><br><a href="URL">Title</a></dt> ... <dd>Authors</dd>
        for m in re.finditer(
                r'<dt class="ptitle">(?:<br>)?<a href="(/content/[^"]+)">([^<]+)</a></dt>'
                r'(.*?)(?=<dt class="ptitle"|</dl>|\Z)', html, re.S):
            url, title, tail = m.group(1), m.group(2).strip(), m.group(3)
            # First <dd> after the title is usually the author list
            am = re.search(r'<dd>\s*([^<]+?)\s*</dd>', tail)
            authors = am.group(1).strip() if am else ''
            out.append({
                'title': re.sub(r'\s+', ' ', title),
                'authors': re.sub(r'\s+', ' ', authors),
                'paper_url': CVF_BASE + url,
            })
        if out:
            break
    return out


def classify_domain(text):
    t = text.lower()
    if any(k in t for k in EXCLUDE_KW):
        return None
    if any(k in t for k in AD_KW):
        return 'AD'
    if any(k in t for k in OD_KW):
        return 'OD'
    if any(k in t for k in CLS_KW):
        return 'CLS'
    return None


def guess_dataset(text):
    t = text.lower()
    for ds in ['mvtec ad 2', 'mvtec loco', 'mvtec 3d', 'mvtec ad', 'visa',
               'mpdd', 'btad', 'real-iad', 'coco', 'lvis', 'pascal voc',
               'imagenet', 'cifar-100', 'cifar-10', 'cub-200', 'duts']:
        if ds in t:
            return ds
    return ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--conf', action='append', default=[],
                    help='Conference: CVPR / ICCV / ECCV / WACV (repeatable)')
    ap.add_argument('--year', action='append', default=[],
                    help='Year, paired positionally with --conf (repeatable)')
    ap.add_argument('--domain', choices=['ad', 'od', 'cls', 'all'], default='all')
    ap.add_argument('--out', default='cvf_results.csv')
    args = ap.parse_args()

    if not args.conf:
        # Default: the most recent proceedings for all four venues
        args.conf = ['CVPR', 'ICCV', 'ECCV', 'WACV']
        args.year = ['2025', '2025', '2024', '2025']
    pairs = list(zip(args.conf, args.year))

    want = None if args.domain == 'all' else args.domain.upper()
    rows = []
    for conf, year in pairs:
        print(f'Fetching {conf} {year} …', file=sys.stderr)
        papers = parse_conference(conf, year)
        print(f'  {len(papers)} papers total', file=sys.stderr)
        for p in papers:
            dom = classify_domain(p['title'])
            if not dom:
                continue
            if want and dom != want:
                continue
            rows.append({
                'domain': dom,
                'conference': conf,
                'year': year,
                'title': p['title'],
                'authors': p['authors'],
                'paper_url': p['paper_url'],
                'dataset_guess': guess_dataset(p['title']),
                'category_guess': '',
                'abstract_snippet': '',
            })
        time.sleep(3)

    with open(args.out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=[
            'domain', 'conference', 'year', 'title', 'authors',
            'paper_url', 'dataset_guess', 'category_guess', 'abstract_snippet'])
        w.writeheader()
        w.writerows(rows)

    by_dom = {}
    for r in rows:
        by_dom[r['domain']] = by_dom.get(r['domain'], 0) + 1
    print(f'\nWrote {len(rows)} rows -> {args.out}', file=sys.stderr)
    print(f'  by domain: {by_dom}', file=sys.stderr)


if __name__ == '__main__':
    main()
