#!/usr/bin/env python3
"""Extract dataset-relevant table snippets from a batch of PDFs."""
import sys, io, pypdf
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

NAMES = sys.argv[1:] if len(sys.argv) > 1 else []
TAGS = ['MVTec AD','MVTec-AD','MVTec','VisA','MPDD','BTAD','MVTec 3D','MVTec-3D','LOCO']


def main():
    for name in NAMES:
        path = f'.pdf_cache/{name}.pdf'
        try:
            r = pypdf.PdfReader(path)
        except Exception as e:
            print(f'### {name}: ERR {e}'); continue
        # Build a tag → page-index map
        hits = {}
        for i, p in enumerate(r.pages):
            t = p.extract_text() or ''
            for tag in TAGS:
                if tag in t:
                    hits.setdefault(tag, []).append(i)
        # Try to find result tables — look for pages with AUROC mentioned
        result_pages = []
        for i, p in enumerate(r.pages):
            t = p.extract_text() or ''
            if 'AUROC' in t.upper() and any(tag in t for tag in TAGS):
                result_pages.append(i)
        print(f'### {name} ({len(r.pages)} pp), result pages: {result_pages[:5]}')
        # Print first 2 result pages
        for i in result_pages[:2]:
            t = r.pages[i].extract_text() or ''
            print(f'  --- p{i} ---')
            # Trim to a max
            print(t[:3000])
            print()
        print()

main()
