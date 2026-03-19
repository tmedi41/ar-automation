"""
run_ar_automation.py
--------------------
Master runner for the daily AR automation workflow.

  Step 1 — ar_collections.py
           Reads exports/ar_aging.csv → writes database/clean_invoices.csv

  Step 2 — generate_collections_emails.py
           Categorises invoices, enforces cooldown, writes email drafts,
           updates database/collections_log.csv, and generates reports.

Run this script once each day instead of the individual scripts.
"""

import os
import re
import sys
import time
import subprocess
from datetime import datetime

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR    = os.path.dirname(SCRIPTS_DIR)

STEP1 = os.path.join(SCRIPTS_DIR, "ar_collections.py")
STEP2 = os.path.join(SCRIPTS_DIR, "generate_collections_emails.py")

SEP      = "=" * 64
THIN_SEP = "-" * 64


# ── Helpers ───────────────────────────────────────────────────────────────────
def _int(text: str, pattern: str, default: int = 0) -> int:
    """Extract the first integer captured by pattern from text."""
    m = re.search(pattern, text)
    return int(m.group(1)) if m else default


def _dollar(text: str, pattern: str, default: str = "$0.00") -> str:
    """Extract the first dollar amount captured by pattern from text."""
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default


def run_step(label: str, script: str) -> tuple[str, float]:
    """
    Run a Python script as a subprocess.
    Prints a live status line, returns (stdout, elapsed_seconds).
    Exits the process if the script fails.
    """
    print(f"  Running {label} ...", end="", flush=True)
    t0     = time.perf_counter()
    result = subprocess.run(
        [sys.executable, script],
        capture_output=True,
        text=True,
        cwd=BASE_DIR,
    )
    elapsed = time.perf_counter() - t0
    print(f"\r  {label:<42} done  ({elapsed:.1f}s)")

    if result.returncode != 0:
        print()
        print(f"  [ERROR] {label} failed (exit code {result.returncode})")
        print(THIN_SEP)
        # Show stderr first, fall back to stdout for scripts that write
        # errors there
        err = result.stderr.strip() or result.stdout.strip()
        for line in err.splitlines():
            print(f"    {line}")
        print(SEP)
        sys.exit(result.returncode)

    return result.stdout, elapsed


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
run_start = time.perf_counter()
now_str   = datetime.now().strftime("%B %d, %Y  %I:%M %p")

print()
print(SEP)
print("  AR AUTOMATION — DAILY RUN")
print(f"  {now_str}")
print(SEP)
print()

# ── Step 1 ───────────────────────────────────────────────────────────────────
out1, t1 = run_step("Step 1 — AR Export Processing", STEP1)

invoices_loaded  = _int(out1, r"Total invoice rows extracted:\s+(\d+)")
unique_customers = _int(out1, r"Unique customers:\s+(\d+)")

# ── Step 2 ───────────────────────────────────────────────────────────────────
out2, t2 = run_step("Step 2 — Collections Email Generation", STEP2)

total_open_ar   = _dollar(out2, r"Total Open AR\s+(\$[\s\d,\.]+)")
total_past_due  = _dollar(out2, r"Total Past Due\s+(\$[\s\d,\.]+)")
customers_pd    = _int(out2, r"Customers Past Due\s+(\d+)")
payments_detected = _int(out2, r"Payments detected \(marked Paid\)\s+(\d+)")
predue_created  = _int(out2, r"Pre-due reminders created\s+(\d+)")
pastdue_created = _int(out2, r"Past-due notices created\s+(\d+)")
escalations     = _int(out2, r"Escalation emails created\s+(\d+)")
skipped         = _int(out2, r"Customers skipped \(cooldown\)\s+(\d+)")
emails_total    = _int(out2, r"Total emails drafted\s+(\d+)")
outlook_drafts  = _int(out2, r"Outlook drafts created\s+(\d+)")
log_entries     = _int(out2, r"New log entries added\s+(\d+)")

total_elapsed = time.perf_counter() - run_start

# ── Final Summary ─────────────────────────────────────────────────────────────
print()
print(SEP)
print("  DAILY RUN SUMMARY")
print(THIN_SEP)
print(f"  {'AR export processed':<38}  {'✓':>4}")
print(f"  {'Invoices loaded':<38}  {invoices_loaded:>4}")
print(f"  {'Unique customers in AR':<38}  {unique_customers:>4}")
print(THIN_SEP)
print(f"  {'Total Open AR':<38}  {total_open_ar.strip():>12}")
print(f"  {'Total Past Due':<38}  {total_past_due.strip():>12}")
print(f"  {'Customers past due':<38}  {customers_pd:>12}")
print(f"  {'Payments detected (marked Paid)':<38}  {payments_detected:>12}")
print(THIN_SEP)
print(f"  {'Customers contacted':<38}  {emails_total:>4}")
print(f"  {'  Pre-due reminders':<38}  {predue_created:>4}")
print(f"  {'  Past-due notices':<38}  {pastdue_created:>4}")
print(f"  {'  Escalations':<38}  {escalations:>4}")
print(f"  {'Customers skipped (cooldown)':<38}  {skipped:>4}")
print(f"  {'Outlook drafts created':<38}  {outlook_drafts:>4}")
print(f"  {'New log entries added':<38}  {log_entries:>4}")
print(THIN_SEP)
print(f"  {'Total run time':<38}  {total_elapsed:.1f}s")
print(SEP)
print()
print("  Top collection priorities generated.")
print("  AR automation run complete.")
print()
