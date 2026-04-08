"""
ar_collections.py
-----------------
Reads exports/ar_aging.csv (QuickBooks AR Aging Detail export),
standardizes columns, cleans balance/date formatting,
groups invoices by customer, and writes database/clean_invoices.csv.
"""

import os
import re
import sys
import pandas as pd

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_FILE  = os.path.join(BASE_DIR, "exports",  "ar_aging.csv")
OUTPUT_FILE = os.path.join(BASE_DIR, "database", "clean_invoices.csv")

# ── DB helpers (imported early so we can restore ar_aging.csv if missing) ──────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from db_utils import get_db_conn, init_tables, write_clean_invoices
    _db_available = bool(os.environ.get("DATABASE_URL", ""))
except Exception:
    def get_db_conn(): return None   # type: ignore[misc]
    _db_available = False


# ── Load ───────────────────────────────────────────────────────────────────────
# On Railway (DATABASE_URL set), ar_aging.csv may not be on disk after a redeploy.
# Restore it from the stored_files table before attempting to read it.
if not os.path.exists(INPUT_FILE):
    if _db_available:
        _conn = get_db_conn()
        if _conn:
            try:
                with _conn.cursor() as _cur:
                    _cur.execute("SELECT content FROM stored_files WHERE key = 'ar_aging'")
                    _stored = _cur.fetchone()
                if _stored:
                    os.makedirs(os.path.dirname(INPUT_FILE), exist_ok=True)
                    with open(INPUT_FILE, "w", encoding="utf-8") as _fh:
                        _fh.write(_stored[0])
                    print("[INFO] ar_aging.csv restored from database.")
                else:
                    print("[ERROR] ar_aging.csv not found on disk or in database. "
                          "Upload it via the dashboard first.")
                    sys.exit(1)
            except Exception as _restore_err:
                print(f"[ERROR] Failed to restore ar_aging.csv from database: {_restore_err}")
                sys.exit(1)
            finally:
                _conn.close()
        else:
            print("[ERROR] DATABASE_URL is set but connection failed. Cannot restore ar_aging.csv.")
            sys.exit(1)
    else:
        print(f"[ERROR] Input file not found: {INPUT_FILE}")
        sys.exit(1)

df = pd.read_csv(INPUT_FILE, dtype=str)

# Normalize column names to strip stray whitespace
df.columns = df.columns.str.strip()

# ── Column rename map (handles QuickBooks header naming variations) ─────────────
rename_map = {
    "Customer full name": "Customer",
    "Customer Full Name": "Customer",
    "customer full name": "Customer",
    "Num":                "Invoice",
    "Date":               "Invoice_Date",
    "Due date":           "Due_Date",
    "Due Date":           "Due_Date",
    "Open balance":       "Balance",
    "Open Balance":       "Balance",
    "Email":              "Email",
}
df = df.rename(columns=rename_map)

# Keep only the six standard columns we care about (drop any extras like "Sent")
keep = ["Customer", "Invoice", "Invoice_Date", "Due_Date", "Balance", "Email"]
df = df[[c for c in keep if c in df.columns]]


# ── Guard: drop any stray summary / section / total rows ──────────────────────
# These can sneak in if the QuickBooks export is regenerated without manual cleanup.
df["Customer"] = df["Customer"].fillna("").str.strip()
df["Balance"]  = df["Balance"].fillna("").str.strip()

is_summary = (
    df["Customer"].str.upper().str.fullmatch("TOTAL")          # grand total row
    | df["Customer"].str.match(r"^Total\s+for", case=False)   # bucket totals
    | df["Customer"].str.match(r"^\d+\s*-\s*\d+\s+days", case=False)  # bucket headers
    | df["Customer"].str.fullmatch(r"CURRENT", case=False)    # CURRENT bucket header
    | (df["Customer"] == "")                                   # blank rows
)
df = df[~is_summary].copy()


# ── Clean Balance ──────────────────────────────────────────────────────────────
# Handles: 1,470.00  |  ($490.43)  |  -9,685.46  |  $298,569.86  (with spaces)
def clean_balance(val: str) -> float | None:
    s = str(val).strip()
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]
    s = re.sub(r"[$,\s]", "", s)
    if s.startswith("-"):
        negative = True
        s = s[1:]
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None

df["Balance"] = df["Balance"].apply(clean_balance)


# ── Parse & reformat dates ─────────────────────────────────────────────────────
for col in ("Invoice_Date", "Due_Date"):
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], format="mixed", dayfirst=False,
                                 errors="coerce").dt.strftime("%m/%d/%Y")


# ── Tidy string columns ────────────────────────────────────────────────────────
df["Invoice"] = df["Invoice"].fillna("").str.strip()
df["Email"]   = df["Email"].fillna("").str.strip()


# ── Sort: group invoices by customer, then by due date ────────────────────────
df = df.sort_values(["Customer", "Due_Date"], na_position="last").reset_index(drop=True)


# ── Save ───────────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
df.to_csv(OUTPUT_FILE, index=False)

# Persist to Postgres so clean_invoices survives Railway redeploys.
if _db_available:
    try:
        init_tables()
        if write_clean_invoices(df):
            print("Postgres: clean_invoices table updated.")
    except Exception as _db_err:
        print(f"[WARN] Postgres write failed (CSV still saved): {_db_err}")
else:
    print("[INFO] DATABASE_URL not set — Postgres write skipped.")


# ── Report ─────────────────────────────────────────────────────────────────────
print(f"Input : {INPUT_FILE}")
print(f"Output: {OUTPUT_FILE}")
print(f"Total invoice rows extracted: {len(df)}")
print(f"Unique customers:             {df['Customer'].nunique()}")
print()

pd.set_option("display.max_colwidth", 42)
pd.set_option("display.width", 150)
print("First 20 rows of cleaned data:")
print(df.head(20).to_string(index=True))
