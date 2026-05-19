#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_metrics.py — download a paper PDF and print the passages that contain
benchmark numbers, so the rows placed into the dataset ranking sheets can be
verified against the paper's own results tables.

Never trust an abstract or a leaderboard cache for numbers — only the paper's
own results table. This tool surfaces those tables; a human (or the agent)
still reads them and decides the final values.

Usage:
    python extract_metrics.py <cvf_html_url | arxiv_abs_url | pdf_url>
    python extract_metrics.py <url> --kw MVTec VisA MPDD AUROC
"""

import argparse, re, ssl, sys, urllib.request

try:
    import certifi
    SSL_CTX = ssl.create_default_context(cafile=certifi.where())
except Exception:
    SSL_CTX = ssl.create_default_context()

UA = 'NTUT-AIL-PaperBot/1.0'
DEFAULT_KW = ['MVTec', 'VisA', 'MPDD', 'BTAD', 'AUROC', 'AUPRO', 'AP ',
              'COCO', 'ImageNet', 'Top-1', 'mAP', 'S-measure']


def fetch_bytes(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as r:
        return r.read()


def resolve_pdf_url(url):
    """Map a CVF html / arXiv abs URL to a direct PDF URL."""
    if url.endswith('.pdf'):
        return url
    if 'openaccess.thecvf.com' in url and '/html/' in url:
        # /content/CVPRxxxx/html/Name_paper.html -> /content/CVPRxxxx/papers/Name_paper.pdf
        return url.replace('/html/', '/papers/').replace('_paper.html', '_paper.pdf')
    m = re.search(r'arxiv\.org/abs/(\d+\.\d+)', url)
    if m:
        return f'https://arxiv.org/pdf/{m.group(1)}'
    return url


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('url')
    ap.add_argument('--kw', nargs='*', default=DEFAULT_KW)
    args = ap.parse_args()

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print('PyMuPDF not installed — run: pip install pymupdf', file=sys.stderr)
        return 1

    pdf_url = resolve_pdf_url(args.url)
    print(f'PDF: {pdf_url}', file=sys.stderr)
    data = fetch_bytes(pdf_url)
    doc = fitz.open(stream=data, filetype='pdf')

    # Title (page 1, first non-trivial line)
    p1 = doc[0].get_text()
    title = next((l.strip() for l in p1.splitlines() if len(l.strip()) > 15), '')
    print(f'\n=== TITLE: {title} ===\n')

    # For each page that mentions a target keyword + a metric-looking number,
    # print (a) structured tables via find_tables() — reliable grid layout —
    # and (b) the raw text as a fallback.
    for i, page in enumerate(doc):
        t = page.get_text()
        if not (any(k in t for k in args.kw) and re.search(r'\b\d{2}\.\d', t)):
            continue
        print(f'====== page {i + 1} ======')
        try:
            tabs = page.find_tables()
            for ti, tab in enumerate(tabs.tables):
                grid = tab.extract()
                if not grid:
                    continue
                print(f'--- TABLE {ti + 1} (page {i + 1}) ---')
                for rrow in grid:
                    cells = [(' '.join(str(c).split()) if c else '') for c in rrow]
                    print(' | '.join(cells))
                print()
        except Exception as e:
            print(f'(table detection failed: {e})')
        # Raw text fallback (kept short)
        print('--- raw text ---')
        print(t.strip())
        print()


if __name__ == '__main__':
    sys.exit(main())
