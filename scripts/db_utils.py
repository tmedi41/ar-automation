"""
db_utils.py
-----------
Shared database helpers for AR Automation scripts.
All scripts import from here so the DB connection and table schema
are defined in exactly one place.

When DATABASE_URL is not set (local dev without Postgres), every
function returns None / False so callers can fall back to CSV.
"""

import os

import pandas as pd
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(_BASE_DIR, ".env"), override=True)

# ── Column name constants (match CSV headers exactly) ────────────────────────

LOG_COLUMNS = [
    "Customer", "Invoice", "Email_Sent_Date", "Followup_Date",
    "Status", "Paid_Date", "Email_Type", "Notes",
]

CLEAN_INV_COLUMNS = ["Customer", "Invoice", "Invoice_Date", "Due_Date", "Balance", "Email"]

INTERACTIONS_COLUMNS = ["Date", "Customer", "Invoice", "Type", "Notes"]


# ── Connection factory ────────────────────────────────────────────────────────

def get_db_conn():
    """Return a psycopg2 connection, or None if DATABASE_URL is not set."""
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        return None
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    return psycopg2.connect(url)


# ── Table initialisation ──────────────────────────────────────────────────────

def init_tables():
    """Create all AR Automation tables if they don't yet exist."""
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
    finally:
        conn.close()


# ── clean_invoices ────────────────────────────────────────────────────────────

def write_clean_invoices(df: pd.DataFrame) -> bool:
    """Truncate clean_invoices and repopulate from df. Returns True on success."""
    conn = get_db_conn()
    if conn is None:
        return False
    try:
        rows = []
        for _, row in df.iterrows():
            balance = row.get("Balance")
            try:
                balance = float(balance) if balance is not None and str(balance) not in ("", "nan", "NaT", "None") else None
            except (TypeError, ValueError):
                balance = None

            def _s(v):
                s = str(v) if v is not None else ""
                return "" if s in ("nan", "NaT", "None") else s

            rows.append((
                _s(row.get("Customer")),
                _s(row.get("Invoice")),
                _s(row.get("Invoice_Date")),
                _s(row.get("Due_Date")),
                balance,
                _s(row.get("Email")),
            ))

        with conn:
            with conn.cursor() as cur:
                cur.execute("TRUNCATE TABLE clean_invoices")
                if rows:
                    psycopg2.extras.execute_values(cur, """
                        INSERT INTO clean_invoices
                            (customer, invoice, invoice_date, due_date, balance, email)
                        VALUES %s
                    """, rows)
        return True
    finally:
        conn.close()


def read_clean_invoices() -> pd.DataFrame | None:
    """Return clean_invoices as a DataFrame, or None if DB is unavailable."""
    conn = get_db_conn()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT customer, invoice, invoice_date, due_date, balance, email
                FROM clean_invoices
            """)
            rows = cur.fetchall()
        if not rows:
            return pd.DataFrame(columns=CLEAN_INV_COLUMNS)
        df = pd.DataFrame(rows, columns=CLEAN_INV_COLUMNS)
        # Normalise: empty strings for None TEXT fields, NaN for None balance
        for col in ("Customer", "Invoice", "Invoice_Date", "Due_Date", "Email"):
            df[col] = df[col].fillna("")
        return df
    finally:
        conn.close()


# ── collections_log ───────────────────────────────────────────────────────────

def read_collections_log() -> pd.DataFrame | None:
    """Return collections_log as a DataFrame (CSV column names), or None."""
    conn = get_db_conn()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT customer, invoice, email_sent_date, followup_date,
                       status, paid_date, email_type, notes
                FROM collections_log
                ORDER BY id
            """)
            rows = cur.fetchall()
        if not rows:
            return pd.DataFrame(columns=LOG_COLUMNS)
        return pd.DataFrame(rows, columns=LOG_COLUMNS).fillna("")
    finally:
        conn.close()


def append_log_rows(rows: list[dict]) -> None:
    """Insert new rows into collections_log."""
    if not rows:
        return
    conn = get_db_conn()
    if conn is None:
        return
    try:
        data = [
            (
                r.get("Customer", ""),
                r.get("Invoice", ""),
                r.get("Email_Sent_Date", ""),
                r.get("Followup_Date", ""),
                r.get("Status", ""),
                r.get("Paid_Date", ""),
                r.get("Email_Type", ""),
                r.get("Notes", ""),
            )
            for r in rows
        ]
        with conn:
            with conn.cursor() as cur:
                psycopg2.extras.execute_values(cur, """
                    INSERT INTO collections_log
                        (customer, invoice, email_sent_date, followup_date,
                         status, paid_date, email_type, notes)
                    VALUES %s
                """, data)
    finally:
        conn.close()


def mark_invoices_paid(invoice_nums: list[str], paid_date: str) -> int:
    """
    Mark all unpaid log entries for the given invoice numbers as Paid.
    Returns the number of rows updated.
    """
    if not invoice_nums:
        return 0
    conn = get_db_conn()
    if conn is None:
        return 0
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("""
                    UPDATE collections_log
                       SET status = 'Paid', paid_date = %s
                     WHERE status != 'Paid'
                       AND invoice = ANY(%s)
                """, (paid_date, invoice_nums))
                return cur.rowcount
    finally:
        conn.close()


# ── customer_interactions ─────────────────────────────────────────────────────

def read_customer_interactions() -> pd.DataFrame | None:
    """Return customer_interactions as a DataFrame, or None if DB unavailable."""
    conn = get_db_conn()
    if conn is None:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT date, customer, invoice, type, notes
                FROM customer_interactions
                ORDER BY id
            """)
            rows = cur.fetchall()
        if not rows:
            return pd.DataFrame(columns=INTERACTIONS_COLUMNS)
        return pd.DataFrame(rows, columns=INTERACTIONS_COLUMNS).fillna("")
    finally:
        conn.close()
