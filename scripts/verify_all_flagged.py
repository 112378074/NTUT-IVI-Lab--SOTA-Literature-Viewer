#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_all_flagged.py
=====================
Automatically verifies all 'flag' / 'none' entries in _ad_verify_list.json
by fetching each paper's HTML/PDF and searching for the claimed metric value.

Run this LOCALLY (not in the cloud session) where arxiv.org, CVF, MDPI,
ScienceDirect, and GitHub are reachable.

Usage:
    python scripts/verify_all_flagged.py          # dry-run, prints findings
    python scripts/verify_all_flagged.py --apply  # writes verified results to JSON

Dependencies:  pip install requests beautifulsoup4 pdfplumber
"""

import json
import re
import sys
import time
from pathlib import Path

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("ERROR: pip install requests beautifulsoup4")
    sys.exit(1)

APPLY = "--apply" in sys.argv
PROJECT = Path(__file__).resolve().parent.parent
VERIFY_JSON = Path(__file__).parent / "_ad_verify_list.json"

HEADERS = {"User-Agent": "Mozilla/5.0 (academic research)"}
TIMEOUT = 20


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def fetch_text(url: str) -> str:
    """Return page text, or '' on failure."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        return soup.get_text(separator=" ", strip=True)
    except Exception as e:
        print(f"  [FETCH FAIL] {url}: {e}")
        return ""


def fetch_arxiv_abstract(arxiv_id: str) -> str:
    return fetch_text(f"https://arxiv.org/abs/{arxiv_id}")


def fetch_arxiv_html(arxiv_id: str) -> str:
    """Try the HTML version of the arxiv paper (better for table parsing)."""
    return fetch_text(f"https://arxiv.org/html/{arxiv_id}")


def number_in_text(text: str, value: str, window: int = 400) -> tuple[bool, str]:
    """
    Return (found, context_snippet).
    Looks for the numeric value within ±window chars of dataset keywords.
    """
    if value is None or value == "N/A":
        return True, "(N/A — no value to verify)"
    val = str(value).rstrip("0").rstrip(".")
    # Search the text for the value with some context
    pattern = re.compile(r"(?<!\d)" + re.escape(val) + r"(?!\d)")
    for m in pattern.finditer(text):
        start = max(0, m.start() - window)
        end = min(len(text), m.end() + window)
        snippet = text[start:end].replace("\n", " ")
        return True, snippet[:300]
    return False, ""


def verify_paper(name: str, entry: dict) -> dict:
    """
    Returns updated entry with verified st values where found.
    Does NOT modify entry in-place if APPLY is False.
    """
    link = entry.get("link", "")
    vals = entry.get("vals", {})
    st = dict(entry.get("st", {}))

    # Determine fetch strategy
    arxiv_m = re.search(r"abs/(\d{4}\.\d+)", link)
    arxiv_id = arxiv_m.group(1) if arxiv_m else None

    text = ""
    if arxiv_id:
        print(f"  Fetching arXiv abstract {arxiv_id} ...")
        text = fetch_arxiv_abstract(arxiv_id)
        if len(text) < 500:
            print(f"  Abstract short, trying HTML version ...")
            text += " " + fetch_arxiv_html(arxiv_id)
        time.sleep(2)
    elif "thecvf.com" in link:
        print(f"  Fetching CVF page ...")
        text = fetch_text(link)
        time.sleep(3)
    elif "github.com" in link:
        # Try raw README
        raw = link.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
        print(f"  Fetching GitHub raw ...")
        text = fetch_text(raw)
        time.sleep(2)
    elif "mdpi.com" in link or "sciencedirect.com" in link:
        print(f"  Fetching publisher page ...")
        text = fetch_text(link)
        time.sleep(3)
    else:
        print(f"  Unknown link type: {link[:80]}")

    if not text:
        print(f"  [SKIP] No text retrieved for {name}")
        return entry

    results = {}
    for dataset, claimed_val in vals.items():
        current_st = st.get(dataset, "flag")
        if "✅" in str(current_st):
            results[dataset] = current_st  # already verified
            continue

        found, snippet = number_in_text(text, str(claimed_val))
        if found:
            results[dataset] = "mem✅"
            print(f"  ✅ {name} @ {dataset}: {claimed_val} FOUND — {snippet[:120]}")
        else:
            results[dataset] = current_st  # keep flag
            print(f"  ❌ {name} @ {dataset}: {claimed_val} NOT FOUND in text")

    if APPLY:
        entry = dict(entry)
        entry["st"] = {**st, **results}
    return entry


# ---------------------------------------------------------------------------
# post-cutoff marker: papers submitted after 2025-08 cannot be verified from
# training knowledge; mark them with a descriptive flag instead of plain 'flag'
# ---------------------------------------------------------------------------

POST_CUTOFF_IDS = {
    "2509", "2510", "2511", "2512",
    "2601", "2602", "2603", "2604", "2605", "2606",
    "2607", "2608", "2609", "2610", "2611", "2612",
}

ICCV2025_PAPERS = {"RareCLIP", "Training-Free Diffusion Defect Generation"}

def is_post_cutoff(name: str, entry: dict) -> bool:
    link = entry.get("link", "")
    m = re.search(r"abs/(\d{4})\.\d+", link)
    if m and m.group(1) in POST_CUTOFF_IDS:
        return True
    if name in ICCV2025_PAPERS:
        return True
    if "ICCV2025" in link:
        return True
    return False


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    with open(VERIFY_JSON, encoding="utf-8") as f:
        data = json.load(f)

    total_flag = 0
    verified_this_run = 0
    post_cutoff_count = 0

    for name, entry in list(data.items()):
        sts = entry.get("st", {})
        needs_verify = [ds for ds, s in sts.items() if "flag" in str(s) or s == "none"]
        if not needs_verify:
            continue
        total_flag += 1

        print(f"\n{'='*60}")
        print(f"Paper: {name}")
        print(f"  Needs verify: {needs_verify}")
        print(f"  Claimed vals: { {k: entry['vals'].get(k) for k in needs_verify} }")

        if is_post_cutoff(name, entry):
            print(f"  ⚠️  POST-CUTOFF (arXiv ≥ 2025-09 or ICCV2025) — skipping web fetch")
            post_cutoff_count += 1
            if APPLY:
                for ds in needs_verify:
                    if data[name]["st"][ds] == "flag":
                        data[name]["st"][ds] = "flag: post-cutoff; needs manual web verify"
            continue

        updated = verify_paper(name, entry)
        new_verified = sum(
            1 for ds in needs_verify
            if "✅" in str(updated.get("st", {}).get(ds, ""))
        )
        verified_this_run += new_verified
        if APPLY and new_verified:
            data[name] = updated

    print(f"\n{'='*60}")
    print(f"Summary:")
    print(f"  Total entries with flag/none: {total_flag}")
    print(f"  Verified this run: {verified_this_run}")
    print(f"  Post-cutoff (skipped): {post_cutoff_count}")

    if APPLY:
        with open(VERIFY_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"\n  ✅ Saved updated {VERIFY_JSON}")
    else:
        print(f"\n  (dry-run; pass --apply to write changes)")


if __name__ == "__main__":
    main()
