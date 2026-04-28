# Auto-Update Pipeline

Scheduled job that fetches the latest AD / OD papers from arXiv every
**Wednesday and Friday at 02:00**, classifies them, updates the two `.xlsx`
source files, regenerates the data inside `index.html`, and pushes to GitHub
(GitHub Pages then redeploys automatically).

## Files

| File | Purpose |
| --- | --- |
| `update_papers.py` | Main script. Fetches arXiv → classifies → appends rows → regenerates JSON → injects HTML → `git push`. |
| `run_update.bat`   | Windows wrapper invoked by Task Scheduler. Logs to `update_log.txt`. |
| `update_log.txt`   | Append-only run log (auto-created on first run). |

## Manual usage

```bash
cd C:\Users\user\Desktop\Mypaper

# 1. Dry run — fetch & classify only, no writes
python scripts\update_papers.py --dry

# 2. Update files but skip git push (good for first verification)
python scripts\update_papers.py --no-push

# 3. Full run (Excel + HTML update, then git push)
python scripts\update_papers.py
```

## How classification works

For each new arXiv paper (last 7 days, not already in workbook):

- **AD category** is decided by keyword matching on title + abstract. Order:
  Diffusion → Normalizing Flow → Data Augmentation → Reconstruction →
  Representation. If nothing matches, defaults to Representation.
- **AD dataset** is detected from text: MVTec AD 2 / LOCO / 3D / AD / VisA /
  MPDD / BTAD. If no match, defaults to MVTec AD.
- **OD category**: Real-Time → RGB Salient → Few-Shot → General OD (default).
- **OD dataset**: matches against the dataset names the workbook already
  contains; defaults to `COCO 2017 val`.

Classification is best-effort. **Metric numbers (I-AUROC, P-AUROC, mAP, FPS,
params, etc.) are left blank** — these cannot be reliably extracted from
arXiv abstracts and need to be filled in manually after a quick read of the
paper.

## Registering the Task Scheduler entries

Run these in an **elevated PowerShell or cmd** (right-click → Run as
administrator). The task will run as the current user, only while you are
logged in. To make it run when logged out, edit the task in `taskschd.msc`
and tick *Run whether user is logged on or not*.

```
schtasks /Create /TN "AD-OD Paper Auto-Update Wed" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D WED /ST 02:00 /F
schtasks /Create /TN "AD-OD Paper Auto-Update Fri" /TR "C:\Users\user\Desktop\Mypaper\scripts\run_update.bat" /SC WEEKLY /D FRI /ST 02:00 /F
```

Verify:

```
schtasks /Query /TN "AD-OD Paper Auto-Update Wed"
schtasks /Query /TN "AD-OD Paper Auto-Update Fri"
```

Run on demand (useful for testing):

```
schtasks /Run /TN "AD-OD Paper Auto-Update Wed"
```

Remove:

```
schtasks /Delete /TN "AD-OD Paper Auto-Update Wed" /F
schtasks /Delete /TN "AD-OD Paper Auto-Update Fri" /F
```

## Git authentication

The script runs `git push` using the credentials cached in the project's
local `.git/config` and Windows Credential Manager. Push the repo manually
once first (so credentials are cached) — after that the script can push
without prompts.

## Failure handling

- If arXiv is unreachable, retries up to 3× with backoff. If all retries
  fail, the run logs the error and exits without touching files.
- If no new papers are found, the script exits without committing.
- If `git push` fails, the run is logged but the local files remain updated.
  Resolve the git issue and rerun (or wait for the next scheduled trigger).
