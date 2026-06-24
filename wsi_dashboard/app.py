"""
wsi_dashboard/app.py
--------------------
Flask web dashboard for the AR Automation system.
Reads reports/, database/, and data/ from the parent AR_Automation directory.
Run:  python3 wsi_dashboard/app.py
"""

import io
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta

import msal
import requests as http_requests
import anthropic
import pandas as pd
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from flask import Flask, jsonify, redirect, render_template, request, session, url_for

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Make scripts/ importable so we can reuse db_utils helpers
sys.path.insert(0, os.path.join(BASE_DIR, "scripts"))
try:
    from db_utils import (
        write_clean_invoices  as _db_write_clean_invoices,
        read_clean_invoices   as _db_read_clean_invoices,
        read_collections_log  as _db_read_collections_log,
    )
except Exception as _db_import_err:
    print(f"[WARN] Could not import db_utils: {_db_import_err}")
    _db_write_clean_invoices = None
    _db_read_clean_invoices  = None
    _db_read_collections_log = None

load_dotenv(os.path.join(BASE_DIR, ".env"), override=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

DATABASE_URL = os.environ.get("DATABASE_URL", "")


# ── Auth helpers ──────────────────────────────────────────────────────────────

def _is_authenticated() -> bool:
    return session.get("authenticated") is True


def _login_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not _is_authenticated():
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


# ── Postgres helpers ──────────────────────────────────────────────────────────

def get_db_conn():
    """Return a psycopg2 connection, or None if DATABASE_URL is not set."""
    if not DATABASE_URL:
        return None
    url = DATABASE_URL
    # Railway uses postgres:// but psycopg2 requires postgresql://
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    return psycopg2.connect(url)


def init_db():
    """Create Postgres tables if absent, then restore ar_aging.csv to disk."""
    conn = get_db_conn()
    if conn is None:
        return
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS customer_interactions (
                        id         SERIAL PRIMARY KEY,
                        date       TEXT NOT NULL DEFAULT '',
                        customer   TEXT NOT NULL DEFAULT '',
                        invoice    TEXT NOT NULL DEFAULT '',
                        type       TEXT NOT NULL DEFAULT 'Customer Reply',
                        notes      TEXT NOT NULL DEFAULT '',
                        created_at TIMESTAMP DEFAULT NOW()
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS collections_log (
                        id              SERIAL PRIMARY KEY,
                        customer        TEXT NOT NULL DEFAULT '',
                        invoice         TEXT NOT NULL DEFAULT '',
                        email_sent_date TEXT NOT NULL DEFAULT '',
                        followup_date   TEXT NOT NULL DEFAULT '',
                        status          TEXT NOT NULL DEFAULT '',
                        paid_date       TEXT NOT NULL DEFAULT '',
                        email_type      TEXT NOT NULL DEFAULT '',
                        notes           TEXT NOT NULL DEFAULT '',
                        created_at      TIMESTAMP DEFAULT NOW()
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS clean_invoices (
                        id           SERIAL PRIMARY KEY,
                        customer     TEXT NOT NULL DEFAULT '',
                        invoice      TEXT NOT NULL DEFAULT '',
                        invoice_date TEXT NOT NULL DEFAULT '',
                        due_date     TEXT NOT NULL DEFAULT '',
                        balance      NUMERIC(14, 2),
                        email        TEXT NOT NULL DEFAULT ''
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS stored_files (
                        key        TEXT PRIMARY KEY,
                        content    TEXT NOT NULL,
                        updated_at TIMESTAMP DEFAULT NOW()
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS recurring_obligations (
                        id                   SERIAL PRIMARY KEY,
                        vendor_name          TEXT NOT NULL,
                        amount               NUMERIC(12, 2) NOT NULL,
                        typical_day_of_month INT NOT NULL,
                        category             TEXT NOT NULL DEFAULT '',
                        is_active            BOOLEAN NOT NULL DEFAULT TRUE
                    )
                """)
                cur.execute("SELECT COUNT(*) FROM recurring_obligations")
                if cur.fetchone()[0] == 0:
                    _seed = [
                        ("T Bank SBA Loan",        24168.00,  1, "Debt"),
                        ("WSI Consulting Fee",      6106.00,  1, "Payroll"),
                        ("WSI Loan Payment",        3454.11,  1, "Debt"),
                        ("Rent",                    6607.00,  1, "Rent"),
                        ("IT Services",              507.00,  1, "Operations"),
                        ("LMS CPA",                 1100.00,  1, "Professional Fees"),
                        ("Health Insurance BCBS",    4244.00, 30, "Benefits"),
                        ("Sun Life (dental/life)",    380.00, 16, "Benefits"),
                        ("PLIC-SBD Insurance",         79.50, 30, "Benefits"),
                        ("Principal-EIS",             396.25,  2, "Benefits"),
                        ("First Insurance",          1765.00,  1, "Insurance"),
                        ("Adrem / Angle Insurance",  3852.00, 16, "Insurance"),
                        ("BancFirst Service Charge",  663.00, 15, "Bank Fees"),
                        ("Merchant Services Fee",     338.00, 15, "Bank Fees"),
                        ("Chase Credit Card",        4384.00, 15, "Credit Card"),
                    ]
                    for vendor, amt, day, cat in _seed:
                        cur.execute(
                            "INSERT INTO recurring_obligations "
                            "(vendor_name, amount, typical_day_of_month, category) "
                            "VALUES (%s, %s, %s, %s)",
                            (vendor, amt, day, cat),
                        )

        # Restore ar_aging.csv from Postgres so automation scripts can find it
        with conn.cursor() as cur:
            cur.execute("SELECT content FROM stored_files WHERE key = 'ar_aging'")
            row = cur.fetchone()
            if row:
                dest = os.path.join(BASE_DIR, "exports", "ar_aging.csv")
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with open(dest, "w", encoding="utf-8") as fh:
                    fh.write(row[0])
    finally:
        conn.close()


# ── Interactions data abstraction ─────────────────────────────────────────────

def _get_interactions_df() -> pd.DataFrame:
    """Return customer_interactions as a DataFrame from Postgres or CSV fallback."""
    _cols = ["Date", "Customer", "Invoice", "Type", "Notes"]
    if DATABASE_URL:
        try:
            conn = get_db_conn()
            if conn:
                try:
                    with conn.cursor() as cur:
                        cur.execute(
                            "SELECT date, customer, invoice, type, notes "
                            "FROM customer_interactions ORDER BY date DESC, id DESC"
                        )
                        rows = cur.fetchall()
                    if not rows:
                        return pd.DataFrame(columns=_cols)
                    return pd.DataFrame(rows, columns=_cols).fillna("")
                finally:
                    conn.close()
        except Exception as e:
            print(f"[ERROR] _get_interactions_df: DB read failed: {e}")
            return pd.DataFrame(columns=_cols)
    # Local dev fallback — read from CSV
    return _read_csv("data/customer_interactions.csv")


def _get_collections_log_df() -> pd.DataFrame:
    """Return collections_log as a DataFrame from Postgres or CSV fallback."""
    _log_cols = ["Customer", "Invoice", "Email_Sent_Date", "Followup_Date",
                 "Status", "Paid_Date", "Email_Type", "Notes", "Balance"]
    if DATABASE_URL:
        try:
            conn = get_db_conn()
            if conn:
                try:
                    with conn.cursor() as cur:
                        cur.execute(
                            "SELECT customer, invoice, email_sent_date, followup_date, "
                            "status, paid_date, email_type, notes, balance "
                            "FROM collections_log ORDER BY id"
                        )
                        rows = cur.fetchall()
                    if not rows:
                        return pd.DataFrame(columns=_log_cols)
                    df = pd.DataFrame(rows, columns=_log_cols)
                    for col in _log_cols[:-1]:  # fill text cols; leave Balance as numeric
                        df[col] = df[col].fillna("")
                    return df
                finally:
                    conn.close()
        except Exception as e:
            print(f"[ERROR] _get_collections_log_df: DB read failed: {e}")
            return pd.DataFrame(columns=_log_cols)
    return _read_csv("database/collections_log.csv")


def _append_interaction(date: str, customer: str, invoice: str, type_: str, notes: str):
    """Insert a new interaction row into Postgres or the CSV fallback."""
    conn = get_db_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "INSERT INTO customer_interactions (date, customer, invoice, type, notes) "
                        "VALUES (%s, %s, %s, %s, %s)",
                        (date, customer, invoice, type_, notes),
                    )
        finally:
            conn.close()
        return

    # CSV fallback
    csv_path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    new_row = pd.DataFrame([{
        "Date": date, "Customer": customer,
        "Invoice": invoice, "Type": type_, "Notes": notes,
    }])
    if os.path.exists(csv_path):
        existing = pd.read_csv(csv_path, dtype=str)
        updated  = pd.concat([existing, new_row], ignore_index=True)
    else:
        updated = new_row
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    updated.to_csv(csv_path, index=False)


def _delete_interaction(raw_date: str, invoice: str, customer: str, notes: str) -> bool:
    """Delete a matching interaction. Returns True if a row was removed."""
    conn = get_db_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "DELETE FROM customer_interactions "
                        "WHERE date=%s AND invoice=%s AND customer=%s AND notes=%s",
                        (raw_date, invoice, customer, notes),
                    )
                    return cur.rowcount > 0
        finally:
            conn.close()

    # CSV fallback
    csv_path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    if not os.path.exists(csv_path):
        return False
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    mask = (
        (df["Date"].str.strip()     == raw_date) &
        (df["Invoice"].str.strip()  == invoice)  &
        (df["Customer"].str.strip() == customer) &
        (df["Notes"].str.strip()    == notes)
    )
    idx = df.index[mask]
    if idx.empty:
        return False
    df = df.drop(idx[0])
    df.to_csv(csv_path, index=False)
    return True


# ── Data helpers ─────────────────────────────────────────────────────────────

def _read_csv(rel_path: str) -> pd.DataFrame:
    path = os.path.join(BASE_DIR, rel_path)
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_csv(path, dtype=str)
    except Exception:
        return pd.DataFrame()


def _process_ar_csv_content(content: str) -> pd.DataFrame:
    """Parse raw ar_aging.csv content into a clean DataFrame (mirrors ar_collections.py logic)."""
    df = pd.read_csv(io.StringIO(content), dtype=str)
    df.columns = df.columns.str.strip()
    rename_map = {
        "Customer full name": "Customer", "Customer Full Name": "Customer",
        "customer full name": "Customer",
        "Num": "Invoice", "Date": "Invoice_Date",
        "Due date": "Due_Date", "Due Date": "Due_Date",
        "Open balance": "Balance", "Open Balance": "Balance",
        "Email": "Email",
    }
    df = df.rename(columns=rename_map)
    keep = ["Customer", "Invoice", "Invoice_Date", "Due_Date", "Balance", "Email"]
    df = df[[c for c in keep if c in df.columns]]
    df["Customer"] = df["Customer"].fillna("").str.strip()
    df["Balance"]  = df["Balance"].fillna("").str.strip()
    is_summary = (
        df["Customer"].str.upper().str.fullmatch("TOTAL")
        | df["Customer"].str.match(r"^Total\s+for", case=False)
        | df["Customer"].str.match(r"^\d+\s*-\s*\d+\s+days", case=False)
        | df["Customer"].str.fullmatch(r"CURRENT", case=False)
        | (df["Customer"] == "")
    )
    df = df[~is_summary].copy()

    def _clean_bal(val: str):
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

    df["Balance"] = df["Balance"].apply(_clean_bal)
    for col in ("Invoice_Date", "Due_Date"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="mixed", dayfirst=False,
                                     errors="coerce").dt.strftime("%m/%d/%Y")
    df["Invoice"] = df["Invoice"].fillna("").str.strip()
    if "Email" in df.columns:
        df["Email"] = df["Email"].fillna("").str.strip()
    return df.sort_values(["Customer", "Due_Date"], na_position="last").reset_index(drop=True)


def get_metrics() -> dict:
    metrics = {
        "total_open_ar":      "$0.00",
        "total_past_due":     "$0.00",
        "payments_this_week": 0,
        "drafts_this_week":   0,
        "generated":          "",
        "customers_past_due": 0,
    }

    today = datetime.today()

    # ── AR totals: read directly from clean_invoices (Postgres) ──────────────
    ci = None
    if DATABASE_URL and _db_read_clean_invoices:
        try:
            ci = _db_read_clean_invoices()
            print(f"[INFO] get_metrics: read {len(ci) if ci is not None else 'None'} rows from clean_invoices")
        except Exception as e:
            print(f"[ERROR] get_metrics: clean_invoices read failed: {e}")
            ci = None

    if ci is not None and not ci.empty:
        ci["Balance"] = pd.to_numeric(ci["Balance"], errors="coerce").fillna(0.0)
        ci["_due"]    = pd.to_datetime(ci["Due_Date"], format="%m/%d/%Y", errors="coerce")
        total_open    = ci["Balance"].sum()
        past_due_mask = ci["_due"].notna() & (ci["_due"] < pd.Timestamp(today.date()))
        past_due_ci   = ci[past_due_mask]
        metrics["total_open_ar"]      = f"${total_open:,.2f}"
        metrics["total_past_due"]     = f"${past_due_ci['Balance'].sum():,.2f}"
        metrics["customers_past_due"] = int(past_due_ci["Customer"].nunique())
        metrics["generated"]          = today.strftime("%B %d, %Y")
        print(f"[INFO] get_metrics: total_open=${total_open:,.2f}, "
              f"past_due_rows={len(past_due_ci)}, "
              f"customers_past_due={metrics['customers_past_due']}")
    elif not DATABASE_URL:
        # Local dev fallback — read from ar_summary.txt if present
        summary_path = os.path.join(BASE_DIR, "reports", "ar_summary.txt")
        if os.path.exists(summary_path):
            text = open(summary_path, encoding="utf-8").read()
            m = re.search(r"Total Open AR\s+\$\s*([\d,]+\.?\d*)", text)
            if m:
                metrics["total_open_ar"] = f"${float(m.group(1).replace(',', '')):,.2f}"
            m = re.search(r"Total Past Due\s+\$\s*([\d,]+\.?\d*)", text)
            if m:
                metrics["total_past_due"] = f"${float(m.group(1).replace(',', '')):,.2f}"
            m = re.search(r"Customers Past Due\s+([\d,]+)", text)
            if m:
                metrics["customers_past_due"] = int(m.group(1).replace(",", ""))
            m = re.search(r"Generated:\s*(.+)", text)
            if m:
                metrics["generated"] = m.group(1).strip()
    else:
        print("[WARN] get_metrics: DATABASE_URL is set but clean_invoices returned empty — "
              "upload ar_aging.csv to populate it")

    # ── Payments / contacts: read directly from collections_log (Postgres) ───
    log = _get_collections_log_df()
    if not log.empty:
        week_start    = today.date() - timedelta(days=today.date().weekday())  # most recent Monday
        week_start_ts = pd.Timestamp(week_start)
        tomorrow_ts   = pd.Timestamp(today.date() + timedelta(days=1))

        log["_paid_dt"] = pd.to_datetime(log.get("Paid_Date", ""),       format="%m/%d/%Y", errors="coerce")
        log["_sent_dt"] = pd.to_datetime(log.get("Email_Sent_Date", ""), format="%m/%d/%Y", errors="coerce")
        has_customer    = log["Customer"].fillna("").str.strip().ne("")
        has_invoice     = log["Invoice"].fillna("").str.strip().ne("")

        paid_mask = (
            (log["Status"].fillna("") == "Paid")
            & log["_paid_dt"].notna()
            & (log["_paid_dt"] >= week_start_ts)
            & (log["_paid_dt"] < tomorrow_ts)
            & has_customer & has_invoice
        )
        sent_mask = (
            log["_sent_dt"].notna()
            & (log["_sent_dt"] >= week_start_ts)
            & (log["_sent_dt"] < tomorrow_ts)
            & has_customer
        )
        metrics["payments_this_week"] = int(paid_mask.sum())
        metrics["drafts_this_week"]   = int(sent_mask.sum())
        print(f"[INFO] get_metrics: log_rows={len(log)}, "
              f"payments_this_week={metrics['payments_this_week']}, "
              f"drafts_this_week={metrics['drafts_this_week']}")
    else:
        print("[INFO] get_metrics: collections_log is empty")

    return metrics


def get_priority_customers() -> list[dict]:
    """Build priority customer list directly from clean_invoices (Postgres-first)."""
    ci = None
    if DATABASE_URL and _db_read_clean_invoices:
        try:
            ci = _db_read_clean_invoices()
            print(f"[INFO] get_priority_customers: read {len(ci) if ci is not None else 'None'} rows")
        except Exception as e:
            print(f"[ERROR] get_priority_customers: DB read failed: {e}")
            ci = None

    if ci is None or ci.empty:
        # Local dev fallback — use the pre-computed summary CSV if present
        summary = _read_csv("reports/collections_summary.csv")
        if summary.empty:
            return []
        summary["Total_Balance"]     = pd.to_numeric(summary["Total_Balance"],     errors="coerce").fillna(0)
        summary["Max_Days_Past_Due"] = pd.to_numeric(summary["Max_Days_Past_Due"], errors="coerce").fillna(0).astype(int)
        summary["Invoice_Count"]     = pd.to_numeric(summary["Invoice_Count"],     errors="coerce").fillna(0).astype(int)
        summary = summary.sort_values("Total_Balance", ascending=False).head(20)
        return [
            {
                "customer":    row["Customer"],
                "balance":     f"${row['Total_Balance']:,.2f}",
                "balance_raw": float(row["Total_Balance"]),
                "invoices":    int(row["Invoice_Count"]),
                "max_dpd":     int(row["Max_Days_Past_Due"]),
                "type":        str(row.get("Email_Type", "")),
            }
            for _, row in summary.iterrows()
        ]

    # ── Compute live from clean_invoices ─────────────────────────────────────
    today = pd.Timestamp.today().normalize()
    ci["Balance"]        = pd.to_numeric(ci["Balance"], errors="coerce").fillna(0.0)
    ci["_due"]           = pd.to_datetime(ci["Due_Date"], format="%m/%d/%Y", errors="coerce")
    ci["Days_Past_Due"]  = (today - ci["_due"]).dt.days.fillna(0).clip(lower=0).astype(int)
    ci["Days_Until_Due"] = (ci["_due"] - today).dt.days.fillna(9999).astype(int)

    _SEVERITY = {"PRE_DUE": 0, "PAST_DUE": 1, "ESCALATION": 2}

    def _categorize(row):
        dpd = row["Days_Past_Due"]
        dud = row["Days_Until_Due"]
        if dpd > 20:         return "ESCALATION"
        if 1 <= dpd <= 20:   return "PAST_DUE"
        if 0 <= dud <= 7:    return "PRE_DUE"
        return None

    ci["Category"] = ci.apply(_categorize, axis=1)

    rows = []
    for customer, grp in ci.groupby("Customer", sort=False):
        total_bal = grp["Balance"].sum()
        if total_bal <= 0:
            continue
        max_dpd   = int(grp["Days_Past_Due"].max())
        cats      = [c for c in grp["Category"].tolist() if c and c in _SEVERITY]
        etype     = max(cats, key=lambda c: _SEVERITY[c]) if cats else ""
        rows.append({
            "customer":    customer,
            "balance":     f"${total_bal:,.2f}",
            "balance_raw": float(total_bal),
            "invoices":    len(grp),
            "max_dpd":     max_dpd,
            "type":        etype,
        })

    rows.sort(key=lambda r: -r["balance_raw"])
    print(f"[INFO] get_priority_customers: returning {len(rows[:20])} customers")
    return rows[:20]


def get_recent_replies() -> list[dict]:
    df = _get_interactions_df()
    if df.empty:
        return []
    df = df.fillna("")
    df = df[df["Customer"].str.strip() != ""]
    df["_dt"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["_dt"]).sort_values("_dt", ascending=False).head(20)

    rows = []
    for _, row in df.iterrows():
        rows.append({
            "date":     row["_dt"].strftime("%m/%d/%Y"),
            "raw_date": row["Date"],
            "invoice":  row.get("Invoice", ""),
            "customer": row["Customer"],
            "notes":    row.get("Notes", "") or "(no note)",
        })
    return rows


def get_customer_replies() -> list:
    df = _get_interactions_df()
    if df.empty:
        return []
    df = df[df["Notes"].str.strip() != ""]
    if df.empty:
        return []
    df["_sort"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values("_sort", ascending=False).drop(columns=["_sort"])
    return df[["Date", "Invoice", "Customer", "Notes"]].rename(
        columns={"Date": "date", "Invoice": "invoice", "Customer": "customer", "Notes": "notes"}
    ).to_dict(orient="records")


def parse_weekly_report() -> dict:
    """Build weekly report data directly from PostgreSQL — never reads from files."""
    today      = datetime.today()
    week_start = today.date() - timedelta(days=today.date().weekday())  # most recent Monday

    result = {
        "generated":         today.strftime("%B %d, %Y"),
        "week_of":           f"{week_start.strftime('%B %-d')} \u2013 {today.strftime('%B %-d, %Y')}",
        "total_contacts":    0,
        "pre_due":           0,
        "past_due_notices":  0,
        "escalations":       0,
        "replies":           [],
        "past_due_accounts": [],
    }

    # ── Metrics from collections_log ──────────────────────────────────────────
    log = _get_collections_log_df()
    if not log.empty:
        week_start_ts = pd.Timestamp(week_start)
        tomorrow_ts   = pd.Timestamp(today.date() + timedelta(days=1))
        log["_sent_dt"] = pd.to_datetime(
            log.get("Email_Sent_Date", ""), format="%m/%d/%Y", errors="coerce"
        )
        week_mask = (
            log["_sent_dt"].notna()
            & (log["_sent_dt"] >= week_start_ts)
            & (log["_sent_dt"] < tomorrow_ts)
        )
        week_log   = log[week_mask]
        email_type = week_log["Email_Type"].str.upper().str.strip()

        result["total_contacts"]   = int(week_log["Customer"].nunique())
        result["pre_due"]          = int((email_type == "PRE_DUE").sum())
        result["past_due_notices"] = int((email_type == "PAST_DUE").sum())
        result["escalations"]      = int((email_type == "ESCALATION").sum())

    # ── Customer replies from customer_interactions ───────────────────────────
    df_int = _get_interactions_df()
    if not df_int.empty:
        df_int = df_int[df_int["Notes"].str.strip() != ""].copy()
        df_int["_dt"] = pd.to_datetime(df_int["Date"], errors="coerce")
        df_int = df_int.dropna(subset=["_dt"]).sort_values("_dt", ascending=False)
        result["replies"] = [
            {
                "date":     row["_dt"].strftime("%m/%d/%Y"),
                "invoice":  row.get("Invoice", ""),
                "customer": row["Customer"],
                "notes":    row.get("Notes", ""),
            }
            for _, row in df_int.iterrows()
        ]

    # ── Past due accounts from clean_invoices ─────────────────────────────────
    ci = None
    if DATABASE_URL and _db_read_clean_invoices:
        try:
            ci = _db_read_clean_invoices()
        except Exception as e:
            print(f"[ERROR] parse_weekly_report: clean_invoices read failed: {e}")

    if ci is not None and not ci.empty:
        ci["Balance"] = pd.to_numeric(ci["Balance"], errors="coerce").fillna(0.0)
        ci["_due"]    = pd.to_datetime(ci["Due_Date"], format="%m/%d/%Y", errors="coerce")
        today_ts      = pd.Timestamp(today.date())
        past_due      = ci[
            ci["_due"].notna() & (ci["_due"] < today_ts) & (ci["Balance"] > 0)
        ].copy()
        past_due["Days_Past_Due"] = (today_ts - past_due["_due"]).dt.days.astype(int)

        accounts = []
        for customer, grp in past_due.groupby("Customer", sort=False):
            total_bal = grp["Balance"].sum()
            if total_bal <= 0:
                continue
            max_dpd  = int(grp["Days_Past_Due"].max())
            invoices = [
                {
                    "number": str(row["Invoice"]),
                    "amount": f"${row['Balance']:,.2f}",
                    "days":   int(row["Days_Past_Due"]),
                }
                for _, row in grp.iterrows()
            ]
            accounts.append({
                "rank":         0,
                "customer":     customer,
                "total":        f"${total_bal:,.2f}",
                "total_raw":    total_bal,
                "days_past_due": max_dpd,
                "invoices":     invoices,
                "last_update":  "",
            })

        accounts.sort(key=lambda a: -a["total_raw"])
        for i, acct in enumerate(accounts[:10], 1):
            acct["rank"] = i
            del acct["total_raw"]
        result["past_due_accounts"] = accounts[:10]

    return result


def enrich_last_updates(past_due_accounts: list) -> list:
    """Override last_update on each past-due account using the most recent
    matching entry in customer_interactions, keyed by invoice number."""
    df = _get_interactions_df()
    if df.empty:
        return past_due_accounts

    df = df[df["Notes"].str.strip() != ""]
    df["_dt"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["_dt"])
    if df.empty:
        return past_due_accounts

    # Build {invoice_number: row_with_max_date}
    latest = (
        df.sort_values("_dt")
          .groupby(df["Invoice"].str.strip())
          .last()
    )

    for acct in past_due_accounts:
        best_dt   = None
        best_text = ""
        for inv in acct.get("invoices", []):
            inv_num = str(inv["number"]).strip()
            if inv_num in latest.index:
                row = latest.loc[inv_num]
                if best_dt is None or row["_dt"] > best_dt:
                    best_dt   = row["_dt"]
                    best_text = f"{row['_dt'].strftime('%m/%d/%Y')} — {row['Notes'].strip()}"
        if best_text:
            acct["last_update"] = best_text

    return past_due_accounts


# ── Payment History data ──────────────────────────────────────────────────────

def get_payment_history_data() -> dict:
    """Build all Payment History tab data from PostgreSQL — no file reads."""
    today = datetime.today()
    month_start_ts = pd.Timestamp(today.replace(day=1).date())
    tomorrow_ts    = pd.Timestamp(today.date() + timedelta(days=1))
    today_ts       = pd.Timestamp(today.date())
    ago7_ts        = pd.Timestamp(today.date() - timedelta(days=7))

    result = {
        "total_collected_month": "$0.00",
        "paid_count_month":      0,
        "past_due_current":      0.0,
        "past_due_7d_ago":       0.0,
        "past_due_trend":        "flat",
        "past_due_delta":        "$0.00",
        "avg_days_to_pay":       "\u2014",
        "escalation_rate":       "0%",
        "weekly_chart":          [],
        "payments":              [],
        "top_overdue":           [],
    }

    log = _get_collections_log_df()
    ci  = None
    if DATABASE_URL and _db_read_clean_invoices:
        try:
            ci = _db_read_clean_invoices()
        except Exception as e:
            print(f"[ERROR] get_payment_history_data: {e}")

    # Pre-compute parsed date columns once
    if not log.empty:
        log = log.copy()
        log["_paid_dt"] = pd.to_datetime(log.get("Paid_Date",       ""), format="%m/%d/%Y", errors="coerce")
        log["_sent_dt"] = pd.to_datetime(log.get("Email_Sent_Date", ""), format="%m/%d/%Y", errors="coerce")
        log["_etype"]   = log["Email_Type"].str.upper().str.strip()

    ci_balance_map: dict = {}
    if ci is not None and not ci.empty:
        ci_balance_map = {
            str(r["Invoice"]).strip(): float(r["Balance"] or 0)
            for _, r in ci.iterrows()
        }

    # Build a fallback balance map from collections_log (balance stored at time of first contact)
    # Use the first non-null balance per invoice (oldest log entry where balance was recorded)
    log_balance_map: dict = {}
    if not log.empty and "Balance" in log.columns:
        for inv, grp in log.groupby("Invoice"):
            valid = pd.to_numeric(grp["Balance"], errors="coerce").dropna()
            if not valid.empty:
                log_balance_map[str(inv).strip()] = float(valid.iloc[0])

    # ── Total Collected This Month ────────────────────────────────────────
    if not log.empty:
        paid_mask = (
            (log["Status"].str.strip() == "Paid")
            & log["_paid_dt"].notna()
            & (log["_paid_dt"] >= month_start_ts)
            & (log["_paid_dt"] < tomorrow_ts)
        )
        paid_this_month = log[paid_mask]
        result["paid_count_month"] = len(paid_this_month)
        total_collected = 0.0
        for inv in paid_this_month["Invoice"]:
            inv_str = str(inv).strip()
            # Prefer current AR balance; fall back to balance recorded in the log
            amount = ci_balance_map.get(inv_str) or log_balance_map.get(inv_str, 0.0)
            total_collected += amount
        result["total_collected_month"] = f"${total_collected:,.2f}"

    # ── Past Due Trend (current vs 7 days ago from clean_invoices) ────────
    if ci is not None and not ci.empty:
        ci_w = ci.copy()
        ci_w["Balance"] = pd.to_numeric(ci_w["Balance"], errors="coerce").fillna(0.0)
        ci_w["_due"]    = pd.to_datetime(ci_w["Due_Date"], format="%m/%d/%Y", errors="coerce")
        current_pd = float(ci_w[ci_w["_due"].notna() & (ci_w["_due"] < today_ts) & (ci_w["Balance"] > 0)]["Balance"].sum())
        ago7_pd    = float(ci_w[ci_w["_due"].notna() & (ci_w["_due"] < ago7_ts)  & (ci_w["Balance"] > 0)]["Balance"].sum())
        delta      = current_pd - ago7_pd
        result["past_due_current"] = current_pd
        result["past_due_7d_ago"]  = ago7_pd
        result["past_due_delta"]   = f"${abs(delta):,.2f}"
        result["past_due_trend"]   = "up" if delta > 0.01 else ("down" if delta < -0.01 else "flat")

    # ── Avg Days to Pay (email_sent_date → paid_date) ─────────────────────
    if not log.empty:
        paid_log = log[(log["Status"].str.strip() == "Paid") & log["_paid_dt"].notna() & log["_sent_dt"].notna()].copy()
        if not paid_log.empty:
            days_series = (paid_log["_paid_dt"] - paid_log["_sent_dt"]).dt.days
            days_series = days_series[days_series >= 0]
            if not days_series.empty:
                result["avg_days_to_pay"] = f"{days_series.mean():.0f}d"

    # ── Escalation Rate ───────────────────────────────────────────────────
    if not log.empty:
        has_customer = log["Customer"].str.strip().ne("")
        total        = int(has_customer.sum())
        escalations  = int((has_customer & (log["_etype"] == "ESCALATION")).sum())
        if total > 0:
            result["escalation_rate"] = f"{escalations / total * 100:.0f}%"

    # ── Weekly Past Due Activity Chart (8 weeks, contacts per week) ───────
    if not log.empty:
        weeks = []
        for i in range(7, -1, -1):
            w_end   = today.date() - timedelta(days=i * 7)
            w_start = w_end - timedelta(days=6)
            mask = (
                log["_sent_dt"].notna()
                & (log["_sent_dt"] >= pd.Timestamp(w_start))
                & (log["_sent_dt"] <  pd.Timestamp(w_end + timedelta(days=1)))
                & log["_etype"].isin(["PAST_DUE", "ESCALATION"])
            )
            weeks.append({
                "label": w_start.strftime("%-m/%-d"),
                "count": int(mask.sum()),
            })
        result["weekly_chart"] = weeks

    # ── Payments Detected (last 20, with best-effort amount) ──────────────
    if not log.empty:
        paid = (
            log[(log["Status"].str.strip() == "Paid") & log["_paid_dt"].notna()]
            .sort_values("_paid_dt", ascending=False)
            .drop_duplicates(subset="Invoice", keep="first")
            .head(20)
        )
        payments = []
        for _, row in paid.iterrows():
            inv = str(row.get("Invoice", "")).strip()
            bal = ci_balance_map.get(inv) or log_balance_map.get(inv) or None
            payments.append({
                "customer":  row.get("Customer", ""),
                "invoice":   inv,
                "amount":    f"${bal:,.2f}" if bal is not None else "\u2014",
                "date_paid": row["_paid_dt"].strftime("%m/%d/%Y"),
            })
        result["payments"] = payments

    # ── Top Overdue Accounts (top 8 by balance, color by type) ───────────
    if ci is not None and not ci.empty:
        ci_w = ci.copy()
        ci_w["Balance"] = pd.to_numeric(ci_w["Balance"], errors="coerce").fillna(0.0)
        ci_w["_due"]    = pd.to_datetime(ci_w["Due_Date"], format="%m/%d/%Y", errors="coerce")
        past_due_ci     = ci_w[ci_w["_due"].notna() & (ci_w["_due"] < today_ts) & (ci_w["Balance"] > 0)]

        # Best email_type per customer from collections_log
        _SEV = {"PRE_DUE": 0, "PAST_DUE": 1, "ESCALATION": 2}
        customer_etype: dict = {}
        if not log.empty:
            for cust, grp in log.groupby("Customer"):
                valid = [t for t in grp["_etype"] if t in _SEV]
                if valid:
                    customer_etype[cust] = max(valid, key=lambda t: _SEV[t])

        overdue = []
        for customer, grp in past_due_ci.groupby("Customer", sort=False):
            total_bal = float(grp["Balance"].sum())
            if total_bal <= 0:
                continue
            overdue.append({
                "customer": customer,
                "balance":  round(total_bal, 2),
                "type":     customer_etype.get(customer, "PAST_DUE"),
            })

        overdue.sort(key=lambda x: -x["balance"])
        result["top_overdue"] = overdue[:8]

    return result


# ── Routes ───────────────────────────────────────────────────────────────────

ADMIN_USERNAME = os.environ.get("DASHBOARD_USERNAME", "TanyaM")
ADMIN_PASSWORD = os.environ.get("DASHBOARD_PASSWORD", "Working$2026")


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "Incorrect username or password."
    return render_template("login.html", error=error)


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@_login_required
def index():
    weekly_data = parse_weekly_report()
    weekly_data["past_due_accounts"] = enrich_last_updates(weekly_data["past_due_accounts"])
    return render_template(
        "index.html",
        metrics=get_metrics(),
        priority=get_priority_customers(),
        replies=get_recent_replies(),
        weekly_data=weekly_data,
        customer_replies=get_customer_replies(),
        payment_data=get_payment_history_data(),
    )


@app.route("/cash-position")
@_login_required
def cash_position():
    obligations = []
    conn = get_db_conn()
    if conn:
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(
                    "SELECT id, vendor_name, amount, typical_day_of_month, "
                    "category, is_active FROM recurring_obligations "
                    "WHERE is_active = TRUE ORDER BY typical_day_of_month, vendor_name"
                )
                obligations = cur.fetchall()
        finally:
            conn.close()

    today = datetime.today()
    current_day = today.day
    total_obligations = sum(float(o["amount"]) for o in obligations)
    already_hit = [o for o in obligations if o["typical_day_of_month"] <= current_day]
    still_pending = [o for o in obligations if o["typical_day_of_month"] > current_day]
    sum_hit = sum(float(o["amount"]) for o in already_hit)
    sum_pending = sum(float(o["amount"]) for o in still_pending)

    return render_template(
        "cash_position.html",
        obligations=obligations,
        already_hit=already_hit,
        still_pending=still_pending,
        total_obligations=f"${total_obligations:,.2f}",
        sum_hit=f"${sum_hit:,.2f}",
        sum_pending=f"${sum_pending:,.2f}",
        current_day=current_day,
        current_month=today.strftime("%B %Y"),
    )


@app.route("/cash-position/calculate", methods=["POST"])
@_login_required
def cash_position_calculate():
    data = request.get_json()
    balance_str = (data.get("balance") or "").strip().replace("$", "").replace(",", "")
    try:
        balance = float(balance_str)
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Please enter a valid dollar amount."})

    obligations = []
    conn = get_db_conn()
    if conn:
        try:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(
                    "SELECT amount, typical_day_of_month FROM recurring_obligations "
                    "WHERE is_active = TRUE"
                )
                obligations = cur.fetchall()
        finally:
            conn.close()

    current_day = datetime.today().day
    total = sum(float(o["amount"]) for o in obligations)
    hit = sum(float(o["amount"]) for o in obligations if o["typical_day_of_month"] <= current_day)
    pending = sum(float(o["amount"]) for o in obligations if o["typical_day_of_month"] > current_day)
    available = balance - pending

    return jsonify({
        "ok": True,
        "entered_balance": f"${balance:,.2f}",
        "total_obligations": f"${total:,.2f}",
        "already_hit": f"${hit:,.2f}",
        "still_pending": f"${pending:,.2f}",
        "available_cash": f"${available:,.2f}",
        "available_raw": available,
    })


@app.route("/run-automation", methods=["POST"])
@_login_required
def run_automation():
    script = os.path.join(BASE_DIR, "scripts", "run_ar_automation.py")
    try:
        result = subprocess.run(
            ["python3", script],
            capture_output=True,
            text=True,
            timeout=180,
            cwd=BASE_DIR,
        )
        ok = result.returncode == 0
        output = (result.stdout or result.stderr or "")[-3000:]
        return jsonify({"ok": ok, "output": output})
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "output": "Automation timed out after 3 minutes."})
    except Exception as exc:
        return jsonify({"ok": False, "output": str(exc)})


@app.route("/upload-ar", methods=["POST"])
@_login_required
def upload_ar():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided."})
    f = request.files["file"]
    if not f.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "File must be a .csv"})

    content = f.read().decode("utf-8", errors="replace")

    # Write to filesystem so automation scripts can read it
    dest = os.path.join(BASE_DIR, "exports", "ar_aging.csv")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, "w", encoding="utf-8") as fh:
        fh.write(content)

    # Persist raw CSV to Postgres so the file survives redeploys
    conn = get_db_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "INSERT INTO stored_files (key, content, updated_at) VALUES (%s, %s, NOW()) "
                        "ON CONFLICT (key) DO UPDATE SET content = EXCLUDED.content, updated_at = NOW()",
                        ("ar_aging", content),
                    )
        finally:
            conn.close()

    # Parse and immediately populate clean_invoices so the table is always current
    if _db_write_clean_invoices is not None and DATABASE_URL:
        try:
            df_clean = _process_ar_csv_content(content)
            if _db_write_clean_invoices(df_clean):
                print(f"[INFO] clean_invoices table populated ({len(df_clean)} rows).")
            else:
                print("[WARN] clean_invoices write returned False — check DATABASE_URL.")
        except Exception as _proc_err:
            print(f"[WARN] Could not populate clean_invoices after upload: {_proc_err}")

    return jsonify({"ok": True})


@app.route("/summarize-reply", methods=["POST"])
@_login_required
def summarize_reply():
    data    = request.get_json()
    invoice = (data.get("invoice") or "").strip()
    reply   = (data.get("reply")   or "").strip()
    if not invoice or not reply:
        return jsonify({"ok": False, "error": "Invoice and reply are required."})

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"ok": False, "error": "ANTHROPIC_API_KEY is not configured."})

    try:
        client  = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=60,
            system=(
                "Summarize customer replies in one phrase, 15 words or fewer. "
                "Capture only the key action or commitment. "
                "CRITICAL: If the reply mentions any dollar amount, you MUST include it verbatim in the summary "
                "— e.g. 'Will pay $1,500 via ACH on April 6' or 'Check #92399 for $2,000 mailed 4/1/26' or "
                "'Payment terms are 60 days, due April 7' or 'Invoice in system to be paid.' "
                "No subjects, no pleasantries, no signatures, no filler. Just the key fact."
            ),
            messages=[{"role": "user", "content": reply}],
        )
        summary = message.content[0].text.strip()
        return jsonify({"ok": True, "summary": summary})
    except anthropic.APIConnectionError:
        return jsonify({"ok": False, "error": "Could not connect to Anthropic API. Check your network or try again."})
    except anthropic.AuthenticationError:
        return jsonify({"ok": False, "error": "Invalid Anthropic API key. Contact your administrator."})
    except anthropic.RateLimitError:
        return jsonify({"ok": False, "error": "Anthropic API rate limit reached. Please wait a moment and try again."})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Summary generation failed: {exc}"})


@app.route("/log-reply", methods=["POST"])
@_login_required
def log_reply():
    data     = request.get_json()
    invoices = (data.get("invoice")  or "").strip()
    customer = (data.get("customer") or "").strip()
    summary  = (data.get("summary")  or "").strip()
    if not invoices or not summary:
        return jsonify({"ok": False, "error": "Invoice and summary are required."})

    today        = datetime.today().strftime("%Y-%m-%d")
    invoice_list = [inv.strip() for inv in invoices.split(",") if inv.strip()]

    # Build per-invoice balance map (only when the AI summary has no dollar amount)
    inv_balance_map: dict = {}
    if "$" not in summary and DATABASE_URL:
        try:
            conn = get_db_conn()
            if conn:
                try:
                    for inv in invoice_list:
                        with conn.cursor() as cur:
                            cur.execute(
                                "SELECT COALESCE(balance, 0) FROM clean_invoices WHERE invoice = %s",
                                (inv,),
                            )
                            row = cur.fetchone()
                            if row and row[0]:
                                inv_balance_map[inv] = float(row[0])
                finally:
                    conn.close()
        except Exception as _bal_err:
            print(f"[WARN] log_reply: balance lookup failed: {_bal_err}")

    for inv in invoice_list:
        inv_summary = summary
        bal = inv_balance_map.get(inv, 0.0)
        if bal > 0:
            inv_summary = f"{summary} \u2014 ${bal:,.2f}"
        _append_interaction(today, customer, inv, "Customer Reply", inv_summary)
    return jsonify({"ok": True, "logged": len(invoice_list)})


@app.route("/admin/fix-reply-balances", methods=["POST"])
@_login_required
def fix_reply_balances():
    """One-time fix: for customer_interactions rows that share the same customer/date/notes
    (meaning they were logged together and received the same combined balance), update each
    row's notes so it shows only that invoice's individual balance from clean_invoices."""
    if not DATABASE_URL:
        return jsonify({"ok": False, "error": "No database configured."})

    fixed = 0
    errors = []
    try:
        conn = get_db_conn()
        if not conn:
            return jsonify({"ok": False, "error": "Could not connect to database."})
        try:
            with conn.cursor() as cur:
                # Find all Customer Reply rows
                cur.execute(
                    "SELECT id, customer, invoice, date, notes "
                    "FROM customer_interactions WHERE type = 'Customer Reply' ORDER BY id"
                )
                rows = cur.fetchall()

            # Group by (customer, date, base_note) where base_note strips trailing " \u2014 $XXX"
            from collections import defaultdict
            groups: dict = defaultdict(list)
            for row_id, customer, invoice, date, notes in rows:
                base = re.sub(r"\s*\u2014\s*\$[\d,]+(?:\.\d+)?\s*$", "", notes or "").strip()
                groups[(customer, date, base)].append((row_id, invoice, notes))

            conn2 = get_db_conn()
            if not conn2:
                return jsonify({"ok": False, "error": "Could not reconnect to database."})
            try:
                for (customer, date, base), entries in groups.items():
                    if len(entries) < 2:
                        continue
                    invoices_in_group = [e[1] for e in entries]
                    # Look up each invoice's individual balance
                    bal_map: dict = {}
                    for inv in invoices_in_group:
                        try:
                            with conn2.cursor() as cur2:
                                cur2.execute(
                                    "SELECT COALESCE(balance, 0) FROM clean_invoices WHERE invoice = %s",
                                    (inv,),
                                )
                                brow = cur2.fetchone()
                                if brow and brow[0]:
                                    bal_map[inv] = float(brow[0])
                        except Exception as _e:
                            errors.append(f"balance lookup {inv}: {_e}")

                    # Update each row with its individual balance
                    for row_id, inv, _old_notes in entries:
                        bal = bal_map.get(inv, 0.0)
                        new_notes = f"{base} \u2014 ${bal:,.2f}" if bal > 0 else base
                        try:
                            with conn2:
                                with conn2.cursor() as cur2:
                                    cur2.execute(
                                        "UPDATE customer_interactions SET notes = %s WHERE id = %s",
                                        (new_notes, row_id),
                                    )
                            fixed += 1
                        except Exception as _e:
                            errors.append(f"update row {row_id}: {_e}")
            finally:
                conn2.close()
        finally:
            conn.close()
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

    return jsonify({"ok": True, "fixed": fixed, "errors": errors})


@app.route("/delete-reply", methods=["POST"])
@_login_required
def delete_reply():
    data     = request.get_json()
    raw_date = (data.get("raw_date") or "").strip()
    invoice  = (data.get("invoice")  or "").strip()
    customer = (data.get("customer") or "").strip()
    notes    = (data.get("notes")    or "").strip()

    if not _delete_interaction(raw_date, invoice, customer, notes):
        return jsonify({"ok": False, "error": "Entry not found."})
    return jsonify({"ok": True})


# ── Weekly Report — Send Report (Graph API draft) ─────────────────────────

def _graph_token() -> str:
    client_id     = os.environ.get("GRAPH_CLIENT_ID", "")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "")
    tenant_id     = os.environ.get("GRAPH_TENANT_ID", "")
    authority     = f"https://login.microsoftonline.com/{tenant_id}"
    app_auth      = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app_auth.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", "MSAL token error"))
    return result["access_token"]


def _build_report_html(weekly_data: dict, customer_replies: list, metrics: dict) -> str:
    _today      = datetime.today()
    _start      = _today - timedelta(days=7)
    week_of     = f"{_start.strftime('%B %d').lstrip('0')} \u2013 {_today.strftime('%B %d, %Y').lstrip('0')}"
    divider     = '<hr style="border:none;border-top:1px solid #e8e8e8;margin:18px 0">'
    cell_style  = "display:inline-block;text-align:left;padding:0 32px 0 0;vertical-align:top"
    label_style = "font-size:11px;color:#999999;text-transform:uppercase;letter-spacing:0.6px;margin:0 0 4px 0"
    value_style = "font-size:20px;font-weight:bold;color:#1a1a1a;margin:0"
    value_red   = "font-size:20px;font-weight:bold;color:#D85A30;margin:0"

    # AR Overview
    ar_overview = f"""
<p style="font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.8px;color:#999999;margin:0 0 12px 0">AR Overview</p>
<div>
  <span style="{cell_style}">
    <p style="{label_style}">Total Open AR</p>
    <p style="{value_style}">{metrics.get('total_open_ar', '—')}</p>
  </span>
  <span style="{cell_style}">
    <p style="{label_style}">Total Past Due</p>
    <p style="{value_red}">{metrics.get('total_past_due', '—')}</p>
  </span>
  <span style="{cell_style}">
    <p style="{label_style}">Customers Past Due</p>
    <p style="{value_style}">{metrics.get('customers_past_due', '—')}</p>
  </span>
</div>"""

    # Contacts This Week
    contacts = f"""
<p style="font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.8px;color:#999999;margin:0 0 12px 0">Contacts This Week</p>
<div>
  <span style="{cell_style}">
    <p style="{label_style}">Total Contacted</p>
    <p style="{value_style}">{weekly_data.get('total_contacts', 0)}</p>
  </span>
  <span style="{cell_style}">
    <p style="{label_style}">Pre-Due Reminders</p>
    <p style="{value_style}">{weekly_data.get('pre_due', 0)}</p>
  </span>
  <span style="{cell_style}">
    <p style="{label_style}">Past-Due Notices</p>
    <p style="{value_style}">{weekly_data.get('past_due_notices', 0)}</p>
  </span>
  <span style="{cell_style}">
    <p style="{label_style}">Escalations</p>
    <p style="{value_style}">{weekly_data.get('escalations', 0)}</p>
  </span>
</div>"""

    # Customer Replies & Notes
    if customer_replies:
        reply_rows = ""
        for r in customer_replies:
            inv_part = f" &middot; inv {r['invoice']}" if r.get("invoice") else ""
            reply_rows += f"""
<div style="padding:8px 0;border-bottom:1px solid #f0f0f0">
  <p style="font-size:13px;font-weight:bold;color:#1a1a1a;margin:0 0 2px 0">{r['customer']}{inv_part}</p>
  <p style="font-size:13px;color:#777777;margin:0">{r['notes']}</p>
</div>"""
        replies_section = f"""
<p style="font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.8px;color:#999999;margin:0 0 4px 0">Customer Replies &amp; Notes</p>
{reply_rows}"""
    else:
        replies_section = """
<p style="font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.8px;color:#999999;margin:0 0 4px 0">Customer Replies &amp; Notes</p>
<p style="font-size:13px;color:#999999;margin:0">No replies logged this week.</p>"""

    return f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background-color:#ffffff">
<div style="max-width:600px;padding:32px 24px;font-family:Arial,sans-serif;color:#1a1a1a">

  <p style="font-size:14px;color:#1a1a1a;margin:0 0 4px 0">Hi Scott,</p>
  <p style="font-size:14px;color:#777777;margin:0 0 20px 0">Here is your AR collections summary for the week of {week_of}.</p>

  {divider}
  {ar_overview}
  {divider}
  {contacts}
  {divider}
  {replies_section}
  {divider}

  <p style="font-size:12px;color:#999999;margin:0">Tanya Medina &middot; Accounting Specialist &middot; Working Solutions Inc.</p>

</div>
</body>
</html>"""


@app.route("/send-report", methods=["POST"])
@_login_required
def send_report():
    sender = os.environ.get("GRAPH_SENDER_EMAIL", "")
    if not sender:
        return jsonify({"ok": False, "error": "GRAPH_SENDER_EMAIL not set in .env"})

    weekly_data = parse_weekly_report()
    weekly_data["past_due_accounts"] = enrich_last_updates(weekly_data["past_due_accounts"])
    customer_replies = get_customer_replies()
    metrics          = get_metrics()
    today_str        = datetime.today().strftime("%B %d, %Y")

    try:
        token = _graph_token()
    except RuntimeError as e:
        return jsonify({"ok": False, "error": str(e)})

    draft = {
        "subject": f"Weekly AR Collections Update — {today_str}",
        "body": {
            "contentType": "HTML",
            "content": _build_report_html(weekly_data, customer_replies, metrics),
        },
        "toRecipients": [
            {"emailAddress": {"address": "scottq@workingsolutions.net"}}
        ],
    }

    url  = f"https://graph.microsoft.com/v1.0/users/{sender}/messages"
    resp = http_requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=draft,
        timeout=15,
    )

    if resp.status_code in (200, 201):
        return jsonify({"ok": True})
    else:
        try:
            err = resp.json().get("error", {}).get("message", resp.text)
        except Exception:
            err = resp.text
        return jsonify({"ok": False, "error": err})


# ── Inbox Scanner ────────────────────────────────────────────────────────────

def _strip_html(html: str) -> str:
    """Remove HTML tags and collapse whitespace to plain text."""
    text = re.sub(r'<[^>]+>', ' ', html)
    return re.sub(r'\s+', ' ', text).strip()


@app.route("/scan-inbox", methods=["POST"])
@_login_required
def scan_inbox():
    sender_email = os.environ.get("GRAPH_SENDER_EMAIL", "")
    if not sender_email:
        return jsonify({"ok": False, "error": "GRAPH_SENDER_EMAIL not set in .env"})

    try:
        token = _graph_token()
    except RuntimeError as e:
        return jsonify({"ok": False, "error": str(e)})

    # ── Build duplicate-check set from existing interactions ──────────────────
    existing = _get_interactions_df()
    logged_keys: set[tuple] = set()
    if not existing.empty:
        for _, row in existing.iterrows():
            c = str(row.get("Customer", "") or "").strip().lower()
            d = str(row.get("Date",     "") or "").strip()[:10]
            if c and d:
                logged_keys.add((c, d))

    # ── Fetch inbox messages from the last 7 days ─────────────────────────────
    since = (datetime.utcnow() - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
    graph_headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.body-content-type="text"',
    }
    base_url = (
        f"https://graph.microsoft.com/v1.0/users/{sender_email}"
        f"/mailFolders/inbox/messages"
    )
    params: dict | None = {
        "$filter": f"receivedDateTime ge {since}",
        "$select": "id,subject,from,receivedDateTime,body",
        "$top":    "50",
    }

    messages: list[dict] = []
    url: str | None = base_url
    for _ in range(5):                           # max 250 messages
        resp = http_requests.get(url, headers=graph_headers, params=params, timeout=30)
        if resp.status_code != 200:
            return jsonify({
                "ok": False,
                "error": f"Graph API error ({resp.status_code}): {resp.text[:300]}",
            })
        page = resp.json()
        messages.extend(page.get("value", []))
        url = page.get("@odata.nextLink")
        params = None                            # nextLink already carries all params
        if not url:
            break

    # ── Subject phrases that identify our outbound AR emails ─────────────────
    TRIGGER_PHRASES = ["Past Due Notice", "Upcoming Invoice Due", "Outstanding Balance"]

    # ── Internal senders to always ignore ────────────────────────────────────
    INTERNAL_SENDERS = {
        "scottq@workingsolutions.net",
        "kimberlys@workingsolutions.net",
        "mikea@workingsolutions.net",
        "rolfes@workingsolutions.net",
        "contact@workingsolutions.net",
        "orders@workingsolutions.net",
    }

    # ── Automated-response detection ──────────────────────────────────────────
    AUTO_SUBJECT_PHRASES = [
        "automatic reply", "auto reply", "out of office",
        "ritm", "ticket", "delivered",
    ]
    AUTO_SENDER_NAME_PHRASES = [
        "automated", "no-reply", "noreply",
        "do not reply", "donotreply", "accounts payable",
    ]
    AUTO_BODY_PHRASES = [
        "this is an automated response",
        "this is an automatic reply",
        "do not reply to this email",
    ]

    # ── Anthropic client (optional — summaries degrade gracefully) ────────────
    api_key   = os.environ.get("ANTHROPIC_API_KEY", "")
    ai_client = anthropic.Anthropic(api_key=api_key) if api_key else None

    logged:           list[dict] = []
    skipped:          int        = 0   # duplicates already logged
    skipped_filtered: int        = 0   # internal / automated

    for msg in messages:
        subject      = (msg.get("subject") or "").strip()
        from_info    = msg.get("from", {}).get("emailAddress", {})
        sender_addr  = (from_info.get("address") or "").strip().lower()
        sender_name  = (from_info.get("name")    or "").strip()
        received_raw = (msg.get("receivedDateTime") or "")

        # If the display name is missing or is itself an email address, derive
        # a readable company name from the sender's domain.
        # e.g. "daniel@excell7.com" → "Excell7"
        #      "ap@bright-horizons.net" → "Bright Horizons"
        if not sender_name or "@" in sender_name:
            _src = sender_name if "@" in sender_name else sender_addr
            _local, _, _domain_full = _src.partition("@")
            _domain_label = _domain_full.split(".")[0] if _domain_full else ""
            # Generic local parts → prefer the domain label as the company name
            _GENERIC = {"info", "contact", "mail", "hello", "support", "sales",
                        "admin", "billing", "accounts", "accounting", "ap", "ar",
                        "office", "reply", "noreply", "no-reply", "donotreply"}
            _base = _domain_label if _local.lower() in _GENERIC else _local
            sender_name = (_base.replace("-", " ").replace("_", " ").replace(".", " ")
                                .title().strip() or _src)
        body_raw     = (msg.get("body", {}).get("content") or "").strip()

        # Skip emails we sent ourselves
        if sender_addr == sender_email.lower():
            continue

        # Skip other internal senders
        if sender_addr in INTERNAL_SENDERS:
            skipped_filtered += 1
            continue

        # Skip automated responses (check subject, sender name, and body)
        subj_lower = subject.lower()
        name_lower = sender_name.lower()
        body_lower = body_raw.lower()[:1000]
        if (
            any(p in subj_lower for p in AUTO_SUBJECT_PHRASES)
            or any(p in name_lower for p in AUTO_SENDER_NAME_PHRASES)
            or any(p in body_lower for p in AUTO_BODY_PHRASES)
        ):
            skipped_filtered += 1
            continue

        # Subject must match at least one trigger phrase
        if not any(p.lower() in subj_lower for p in TRIGGER_PHRASES):
            continue

        # Parse received date → YYYY-MM-DD
        try:
            date_str = datetime.strptime(received_raw[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            date_str = datetime.today().strftime("%Y-%m-%d")

        # Duplicate check: same sender name + same calendar day
        dup_key = (sender_name.lower(), date_str)
        if dup_key in logged_keys:
            skipped += 1
            continue

        # Extract invoice number from subject (last 4-7 digit run, best-effort)
        inv_matches = re.findall(r'\b(\d{4,7})\b', subject)
        invoice = inv_matches[-1] if inv_matches else ""

        # Normalise body to plain text
        body_text = _strip_html(body_raw) if "<" in body_raw else body_raw
        body_text = body_text[:2000]

        # Generate AI summary (≤12 words); fall back to truncated body text
        summary = ""
        if ai_client and body_text:
            try:
                ai_resp = ai_client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=40,
                    system=(
                        "Summarize this customer email reply in one phrase, 12 words or fewer. "
                        "Capture only the key action or commitment — e.g. 'Will pay via ACH Friday' "
                        "or 'Invoice already sent to AP, processing next week'. "
                        "No subjects, no pleasantries, no filler."
                    ),
                    messages=[{"role": "user", "content": body_text}],
                )
                summary = ai_resp.content[0].text.strip()
            except Exception:
                summary = body_text[:120]
        elif body_text:
            summary = body_text[:120]

        # Persist and mark as seen
        _append_interaction(date_str, sender_name, invoice, "Customer Reply", summary)
        logged_keys.add(dup_key)

        logged.append({
            "customer": sender_name,
            "invoice":  invoice,
            "date":     date_str,
            "subject":  subject,
            "summary":  summary,
        })

    return jsonify({
        "ok":              True,
        "new_count":       len(logged),
        "skipped":         skipped,
        "skipped_filtered": skipped_filtered,
        "items":           logged,
    })


# ── Startup ───────────────────────────────────────────────────────────────────

# Run on import so gunicorn workers also initialise the database
init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
