"""
generate_collections_emails.py
-------------------------------
Reads database/clean_invoices.csv and categorises every invoice into one
of three actionable buckets, then generates one email draft per customer
(subject to a 7-day cooldown), a collections summary CSV, and an
executive AR summary report.

Category rules (per invoice):
  PRE_DUE     Days_Until_Due  0 – 7   →  Upcoming reminder
  PAST_DUE    Days_Past_Due   1 – 20  →  Past-due notice
  ESCALATION  Days_Past_Due  > 20     →  Urgent escalation

Email tone follows the most severe category present for that customer.
Invoices outside all three windows are excluded from emails.
"""

import os
import re
import sys
import textwrap
import pandas as pd

# ── Postgres helpers (db_utils lives in the same scripts/ directory) ──────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import db_utils as _db
    _db.init_tables()
    _db_available = bool(os.environ.get("DATABASE_URL", ""))
except Exception:
    _db_available = False

# ── Microsoft Graph API authentication ───────────────────────────────────────
# Credentials are read from the .env file in the project root; they are never
# hard-coded here.  Install dependencies once:
#   pip install msal requests python-dotenv
try:
    from dotenv import load_dotenv
    import msal
    import requests as _requests
    load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))
    _GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID",     "")
    _GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID",     "")
    _GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")
    _GRAPH_SENDER_EMAIL  = os.getenv("GRAPH_SENDER_EMAIL",  "")
    GRAPH_AVAILABLE = bool(_GRAPH_CLIENT_ID and _GRAPH_TENANT_ID
                           and _GRAPH_CLIENT_SECRET and _GRAPH_SENDER_EMAIL)
except ImportError:
    GRAPH_AVAILABLE = False

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_FILE   = os.path.join(BASE_DIR, "database", "clean_invoices.csv")
LOG_FILE     = os.path.join(BASE_DIR, "database", "collections_log.csv")
EMAIL_DIR           = os.path.join(BASE_DIR, "reports", "email_drafts")
SUMMARY_FILE        = os.path.join(BASE_DIR, "reports", "collections_summary.csv")
AR_REPORT           = os.path.join(BASE_DIR, "reports", "ar_summary.txt")
PRIORITY_REPORT     = os.path.join(BASE_DIR, "reports", "collections_priority.txt")
WEEKLY_REPORT       = os.path.join(BASE_DIR, "reports", "weekly_collections_update.txt")
INTERACTIONS_FILE   = os.path.join(BASE_DIR, "data",    "customer_interactions.csv")

LOG_COLUMNS   = ["Customer", "Invoice", "Email_Sent_Date", "Followup_Date",
                 "Status", "Paid_Date", "Email_Type", "Notes"]
COOLDOWN_DAYS = 7
# Severity order: higher index = more urgent
SEVERITY      = {"PRE_DUE": 0, "PAST_DUE": 1, "ESCALATION": 2}

for d in (os.path.join(BASE_DIR, "database"),
          os.path.join(BASE_DIR, "data"),
          os.path.join(BASE_DIR, "reports"),
          EMAIL_DIR):
    os.makedirs(d, exist_ok=True)


# ════════════════════════════════════════════════════════════════════════════
# 1.  LOAD & CATEGORISE INVOICE DATA
# ════════════════════════════════════════════════════════════════════════════
if _db_available:
    _ci = _db.read_clean_invoices()
    if _ci is not None and not _ci.empty:
        df = _ci
    elif os.path.exists(INPUT_FILE):
        # DB table empty — use local CSV fallback (development only)
        df = pd.read_csv(INPUT_FILE, dtype=str)
    else:
        print("[ERROR] clean_invoices table is empty and no CSV fallback found. "
              "Upload ar_aging.csv via the dashboard first.")
        sys.exit(1)
else:
    if os.path.exists(INPUT_FILE):
        df = pd.read_csv(INPUT_FILE, dtype=str)
    else:
        print(f"[ERROR] Input file not found: {INPUT_FILE}")
        sys.exit(1)
df["Balance"]      = pd.to_numeric(df["Balance"], errors="coerce").fillna(0.0)
df["Due_Date"]     = pd.to_datetime(df["Due_Date"],     format="%m/%d/%Y", errors="coerce")
df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], format="%m/%d/%Y", errors="coerce")
df["Invoice"]      = df["Invoice"].fillna("").str.strip()

today = pd.Timestamp.today().normalize()

df["Days_Past_Due"]  = (today - df["Due_Date"]).dt.days   # positive = overdue
df["Days_Until_Due"] = (df["Due_Date"] - today).dt.days   # positive = future


def _categorize(row) -> str | None:
    dud = row["Days_Until_Due"]
    dpd = row["Days_Past_Due"]
    if 0 <= dud <= 7:
        return "PRE_DUE"
    if 1 <= dpd <= 20:
        return "PAST_DUE"
    if dpd > 20:
        return "ESCALATION"
    return None   # not actionable (future > 7 days, or exact-day edge already covered)


df["Category"] = df.apply(_categorize, axis=1)

# Actionable = any invoice needing an email
actionable = df[df["Category"].notna()].copy()

# Keep a strict past-due view for the AR overview numbers
past_due = df[df["Days_Past_Due"] > 0].copy()


# ════════════════════════════════════════════════════════════════════════════
# 2.  LOAD / INITIALISE COLLECTIONS LOG
# ════════════════════════════════════════════════════════════════════════════
if _db_available:
    _log_raw_db = _db.read_collections_log()
    log = _log_raw_db if _log_raw_db is not None else pd.DataFrame(columns=LOG_COLUMNS)
    if log.empty:
        log["Email_Sent_Date"] = pd.NaT
    else:
        log["Email_Sent_Date"] = pd.to_datetime(
            log["Email_Sent_Date"], format="%m/%d/%Y", errors="coerce"
        )
elif os.path.exists(LOG_FILE):
    log = pd.read_csv(LOG_FILE, dtype=str)
    for col in LOG_COLUMNS:
        if col not in log.columns:
            log[col] = ""
    log["Email_Sent_Date"] = pd.to_datetime(
        log["Email_Sent_Date"], format="%m/%d/%Y", errors="coerce"
    )
else:
    log = pd.DataFrame(columns=LOG_COLUMNS)
    log["Email_Sent_Date"] = pd.NaT
    print(f"[INFO] No collections log found — creating {LOG_FILE}")

log["Invoice"] = log["Invoice"].fillna("").str.strip()

# Fast lookup: invoice → most-recent Email_Sent_Date
recent_contact = (
    log.dropna(subset=["Email_Sent_Date"])
       .sort_values("Email_Sent_Date")
       .drop_duplicates(subset=["Invoice"], keep="last")
       .set_index("Invoice")["Email_Sent_Date"]
)

# ── Payment detection ────────────────────────────────────────────────────────
# Any invoice that was logged (Status != "Paid") but is no longer in the
# current AR export is assumed paid/cleared.
payments_detected = 0
if not log.empty:
    _current_inv  = set(df["Invoice"].astype(str).str.strip())
    _unpaid_mask  = log["Status"].fillna("").ne("Paid")
    _missing_mask = ~log["Invoice"].fillna("").str.strip().isin(_current_inv)
    _paid_update  = _unpaid_mask & _missing_mask
    payments_detected = int(_paid_update.sum())
    if payments_detected > 0:
        _paid_date_str = today.strftime("%m/%d/%Y")
        _invoices_to_mark = log.loc[_paid_update, "Invoice"].tolist()
        if _db_available:
            _db.mark_invoices_paid(_invoices_to_mark, _paid_date_str)
        else:
            # CSV fallback: reload, update, and resave
            if os.path.exists(LOG_FILE):
                _csv_log = pd.read_csv(LOG_FILE, dtype=str)
                for _col in LOG_COLUMNS:
                    if _col not in _csv_log.columns:
                        _csv_log[_col] = ""
                _csv_unpaid  = _csv_log["Status"].fillna("").ne("Paid")
                _csv_missing = ~_csv_log["Invoice"].fillna("").str.strip().isin(_current_inv)
                _csv_upd     = _csv_unpaid & _csv_missing
                _csv_log.loc[_csv_upd, "Status"]    = "Paid"
                _csv_log.loc[_csv_upd, "Paid_Date"] = _paid_date_str
                _csv_log.to_csv(LOG_FILE, index=False)
        # Keep in-memory log in sync so sections 10A/10C see paid status
        log.loc[_paid_update, "Status"]    = "Paid"
        log.loc[_paid_update, "Paid_Date"] = _paid_date_str
        print(f"[INFO] {payments_detected} invoice(s) marked as Paid in collections log.")


# ════════════════════════════════════════════════════════════════════════════
# 3.  HELPERS
# ════════════════════════════════════════════════════════════════════════════
def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\s-]", "", name.strip())
    return re.sub(r"\s+", "_", name).strip("_").lower()


def is_recently_contacted(invoice_num: str) -> bool:
    """True if this invoice was emailed within the 7-day cooldown window."""
    inv = str(invoice_num).strip()
    if inv not in recent_contact.index:
        return False
    return (today - recent_contact[inv]).days < COOLDOWN_DAYS


def worst_category(categories) -> str:
    """Return the most severe category from a list."""
    return max(categories, key=lambda c: SEVERITY[c])


def _invoice_section(group: pd.DataFrame, days_col: str, days_label: str) -> str:
    """Format a block of invoice rows with the appropriate days column."""
    header  = (f"  {'Invoice':<12} {'Invoice Date':<16} {'Due Date':<14}"
               f" {days_label:<18} {'Balance':>12}")
    divider = (f"  {'-'*12} {'-'*16} {'-'*14} {'-'*18} {'-'*12}")
    rows = [header, divider]
    for _, row in group.sort_values("Due_Date").iterrows():
        inv_date = row["Invoice_Date"].strftime("%m/%d/%Y") if pd.notna(row["Invoice_Date"]) else "N/A"
        due_date = row["Due_Date"].strftime("%m/%d/%Y")     if pd.notna(row["Due_Date"])     else "N/A"
        days_val = max(int(row[days_col]), 0)   # clamp negatives (edge cases)
        rows.append(
            f"  {str(row['Invoice']):<12} {inv_date:<16} {due_date:<14}"
            f" {days_val:<18} ${row['Balance']:>11,.2f}"
        )
    return "\n".join(rows)


def build_email(customer: str, group: pd.DataFrame,
                total_balance: float, email_type: str) -> str:
    """
    Compose a tiered collections email ready to paste into Outlook.
    Past-due and upcoming invoices appear in clearly labelled sections.
    """
    overdue_grp = group[group["Category"].isin(["PAST_DUE", "ESCALATION"])]
    predue_grp  = group[group["Category"] == "PRE_DUE"]

    # ── Build labelled invoice sections ──────────────────────────────────
    sections = []
    if not overdue_grp.empty:
        sections.append("  Past-Due Invoices:")
        sections.append(_invoice_section(overdue_grp, "Days_Past_Due",  "Days Past Due"))
        sections.append("")
    if not predue_grp.empty:
        label = "  Invoices Due Within 7 Days:" if not overdue_grp.empty else "  Upcoming Invoices (Due Within 7 Days):"
        sections.append(label)
        sections.append(_invoice_section(predue_grp,  "Days_Until_Due", "Days Until Due"))
        sections.append("")

    invoice_block = "\n".join(sections)
    balance_line  = f"Total Outstanding Balance: ${total_balance:,.2f}"

    # ── PRE_DUE ───────────────────────────────────────────────────────────
    if email_type == "PRE_DUE":
        subject = f"Upcoming Invoice Due \u2013 Friendly Reminder \u2013 {customer}"
        body = (
            f"Hello,\n"
            f"\n"
            f"This is a friendly reminder that the following invoice(s) on your "
            f"account will be coming due within the next 7 days. To ensure "
            f"uninterrupted service and avoid any late fees, please plan for "
            f"timely payment.\n"
            f"\n"
            f"{invoice_block}"
            f"\n{balance_line}\n"
            f"\n"
            f"If payment has already been arranged, please disregard this notice. "
            f"Should you have any questions or need to discuss payment terms, "
            f"please don't hesitate to reach out. We appreciate your continued "
            f"business.\n"
            f"\n"
            f"Thank you,\n"
            f"Working Solutions Inc.\n"
        )

    # ── PAST_DUE ──────────────────────────────────────────────────────────
    elif email_type == "PAST_DUE":
        subject = f"Past Due Notice \u2013 {customer}"
        has_upcoming = not predue_grp.empty
        upcoming_note = (
            f" Additionally, please be aware that you have invoice(s) "
            f"coming due in the next 7 days, also listed below."
            if has_upcoming else ""
        )
        body = (
            f"Hello,\n"
            f"\n"
            f"We're reaching out regarding an outstanding balance on your account. "
            f"Please review the invoice(s) below and let us know your expected "
            f"payment date at your earliest convenience.{upcoming_note}\n"
            f"\n"
            f"{invoice_block}"
            f"\n{balance_line}\n"
            f"\n"
            f"If payment has already been sent, please disregard this notice. "
            f"Otherwise, we'd appreciate a quick reply with your expected payment "
            f"date or any questions you may have.\n"
            f"\n"
            f"Thank you,\n"
            f"Working Solutions Inc.\n"
        )

    # ── ESCALATION ────────────────────────────────────────────────────────
    else:
        subject = f"Urgent \u2013 Outstanding Balance \u2013 {customer}"
        has_upcoming = not predue_grp.empty
        upcoming_note = (
            f" Please also note that you have additional invoice(s) "
            f"coming due within 7 days, listed below."
            if has_upcoming else ""
        )
        body = (
            f"Hello,\n"
            f"\n"
            f"This is an urgent notice regarding a past-due balance on your account "
            f"that requires your immediate attention. Despite previous communications, "
            f"the invoices listed below remain unpaid and require your immediate "
            f"attention.{upcoming_note}\n"
            f"\n"
            f"{invoice_block}"
            f"\n{balance_line}\n"
            f"\n"
            f"Please remit payment in full or contact us today to arrange an "
            f"immediate resolution. Failure to respond may result in escalation "
            f"to our formal collections process, which could affect your account "
            f"standing with Working Solutions Inc.\n"
            f"\n"
            f"We strongly encourage you to reach out today so this matter can be "
            f"resolved promptly and without further action.\n"
            f"\n"
            f"Thank you,\n"
            f"Working Solutions Inc.\n"
        )

    return f"Subject: {subject}\n\n{body}"


def _invoice_section_html(group: pd.DataFrame, days_col: str, days_label: str) -> str:
    """Return an HTML table for a group of invoices.
    Invoice number, Due Date, and Balance cells are bolded per spec."""
    th = "style=\"padding:4px 12px 4px 0;text-align:left;border-bottom:1px solid #ccc\""
    td_l = "style=\"padding:3px 12px 3px 0;text-align:left\""
    td_r = "style=\"padding:3px 0;text-align:right\""
    rows = [
        f'<table style="border-collapse:collapse;font-family:monospace;font-size:13px">',
        f'<tr>'
        f'<th {th}>Invoice</th>'
        f'<th {th}>Invoice Date</th>'
        f'<th {th}>Due Date</th>'
        f'<th {th}>{days_label}</th>'
        f'<th {th} style="text-align:right">Balance</th>'
        f'</tr>',
    ]
    for _, row in group.sort_values("Due_Date").iterrows():
        inv_date = row["Invoice_Date"].strftime("%m/%d/%Y") if pd.notna(row["Invoice_Date"]) else "N/A"
        due_date = row["Due_Date"].strftime("%m/%d/%Y")     if pd.notna(row["Due_Date"])     else "N/A"
        days_val = max(int(row[days_col]), 0)
        rows.append(
            f'<tr>'
            f'<td {td_l}><b>{row["Invoice"]}</b></td>'
            f'<td {td_l}>{inv_date}</td>'
            f'<td {td_l}><b>{due_date}</b></td>'
            f'<td {td_l}>{days_val}</td>'
            f'<td {td_r}><b>${row["Balance"]:,.2f}</b></td>'
            f'</tr>'
        )
    rows.append('</table>')
    return "\n".join(rows)


def build_email_html(customer: str, group: pd.DataFrame,
                     total_balance: float, email_type: str) -> tuple[str, str]:
    """Return (subject, html_body) for Outlook drafts via Graph API.
    Mirrors build_email() but produces HTML with bolded invoice numbers,
    due dates, balance cells, and the total balance line."""
    overdue_grp = group[group["Category"].isin(["PAST_DUE", "ESCALATION"])]
    predue_grp  = group[group["Category"] == "PRE_DUE"]

    # ── Build labelled HTML invoice section(s) ───────────────────────────
    sections = []
    if not overdue_grp.empty:
        sections.append("<p><b>Past-Due Invoices:</b></p>")
        sections.append(_invoice_section_html(overdue_grp, "Days_Past_Due", "Days Past Due"))
    if not predue_grp.empty:
        label = "Invoices Due Within 7 Days:" if not overdue_grp.empty else "Upcoming Invoices (Due Within 7 Days):"
        sections.append(f"<p><b>{label}</b></p>")
        sections.append(_invoice_section_html(predue_grp, "Days_Until_Due", "Days Until Due"))

    invoice_block_html = "\n".join(sections)
    balance_line_html  = f"<p><b>Total Outstanding Balance: ${total_balance:,.2f}</b></p>"

    # ── PRE_DUE ──────────────────────────────────────────────────────────
    if email_type == "PRE_DUE":
        subject = f"Upcoming Invoice Due \u2013 Friendly Reminder \u2013 {customer}"
        html_body = (
            f"<p>Hello,</p>"
            f"<p>This is a friendly reminder that the following invoice(s) on your "
            f"account will be coming due within the next 7 days. To ensure "
            f"uninterrupted service and avoid any late fees, please plan for "
            f"timely payment.</p>"
            f"{invoice_block_html}"
            f"{balance_line_html}"
            f"<p>If payment has already been arranged, please disregard this notice. "
            f"Should you have any questions or need to discuss payment terms, "
            f"please don't hesitate to reach out. We appreciate your continued "
            f"business.</p>"
            f"<p>Thank you,<br>Working Solutions Inc.</p>"
        )

    # ── PAST_DUE ─────────────────────────────────────────────────────────
    elif email_type == "PAST_DUE":
        subject = f"Past Due Notice \u2013 {customer}"
        has_upcoming = not predue_grp.empty
        upcoming_note = (
            f" Additionally, please be aware that you have invoice(s) "
            f"coming due in the next 7 days, also listed below."
            if has_upcoming else ""
        )
        html_body = (
            f"<p>Hello,</p>"
            f"<p>We're reaching out regarding an outstanding balance on your account. "
            f"Please review the invoice(s) below and let us know your expected "
            f"payment date at your earliest convenience.{upcoming_note}</p>"
            f"{invoice_block_html}"
            f"{balance_line_html}"
            f"<p>If payment has already been sent, please disregard this notice. "
            f"Otherwise, we'd appreciate a quick reply with your expected payment "
            f"date or any questions you may have.</p>"
            f"<p>Thank you,<br>Working Solutions Inc.</p>"
        )

    # ── ESCALATION ───────────────────────────────────────────────────────
    else:
        subject = f"Urgent \u2013 Outstanding Balance \u2013 {customer}"
        has_upcoming = not predue_grp.empty
        upcoming_note = (
            f" Please also note that you have additional invoice(s) "
            f"coming due within 7 days, listed below."
            if has_upcoming else ""
        )
        html_body = (
            f"<p>Hello,</p>"
            f"<p>This is an urgent notice regarding a past-due balance on your account "
            f"that requires your immediate attention. Despite previous communications, "
            f"the invoices listed below remain unpaid and require your immediate "
            f"attention.{upcoming_note}</p>"
            f"{invoice_block_html}"
            f"{balance_line_html}"
            f"<p>Please remit payment in full or contact us today to arrange an "
            f"immediate resolution. Failure to respond may result in escalation "
            f"to our formal collections process, which could affect your account "
            f"standing with Working Solutions Inc.</p>"
            f"<p>We strongly encourage you to reach out today so this matter can be "
            f"resolved promptly and without further action.</p>"
            f"<p>Thank you,<br>Working Solutions Inc.</p>"
        )

    return subject, html_body


def _parse_email_text(email_text: str) -> tuple[str, str]:
    """
    Split the canonical 'Subject: …\n\n{body}' string produced by
    build_email() into a (subject, body) tuple for use with Outlook.
    """
    subject_line, _, body = email_text.partition("\n\n")
    return subject_line.removeprefix("Subject: "), body


def _split_addresses(raw: str) -> tuple[str, str]:
    """
    Given a comma-separated address string from the CSV Email column,
    return (to_addr, cc_addr) where cc_addr uses Outlook's '; ' separator.
    Empty strings are returned safely when no address is present.
    """
    parts = [a.strip() for a in raw.split(",") if a.strip()]
    to_addr = parts[0] if parts else ""
    cc_addr = "; ".join(parts[1:])
    return to_addr, cc_addr


# Module-level Graph API token cache (acquired once per run)
_graph_token: str | None = None


def _get_graph_token() -> str | None:
    """Acquire (and cache) a Graph API access token via MSAL client credentials."""
    global _graph_token
    if not GRAPH_AVAILABLE:
        return None
    if _graph_token is not None:
        return _graph_token
    try:
        app = msal.ConfidentialClientApplication(
            _GRAPH_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{_GRAPH_TENANT_ID}",
            client_credential=_GRAPH_CLIENT_SECRET,
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        _graph_token = result.get("access_token")
        return _graph_token
    except Exception:
        return None


def create_outlook_draft(subject: str, html_body: str, raw_addresses: str) -> bool:
    """
    Create a saved HTML draft in the sender's Outlook Drafts folder via the
    Microsoft Graph API.  Returns True on success, False otherwise.
    The .txt fallback is always written regardless of this return value.
    """
    token = _get_graph_token()
    if token is None:
        return False
    try:
        to_addr, cc_addr = _split_addresses(raw_addresses)

        to_recipients = [{"emailAddress": {"address": to_addr}}] if to_addr else []
        cc_recipients = (
            [{"emailAddress": {"address": a.strip()}}
             for a in cc_addr.split(";") if a.strip()]
            if cc_addr else []
        )

        payload = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
        }
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }
        url = f"https://graph.microsoft.com/v1.0/users/{_GRAPH_SENDER_EMAIL}/messages"
        resp = _requests.post(url, json=payload, headers=headers, timeout=30)
        return resp.status_code in (200, 201)
    except Exception:
        return False


# ════════════════════════════════════════════════════════════════════════════
# 4.  GENERATE EMAILS  +  UPDATE LOG
# ════════════════════════════════════════════════════════════════════════════
summary_rows    = []
new_log_rows    = []
emails_written  = 0
emails_skipped  = 0
outlook_drafts  = 0
type_counts     = {"PRE_DUE": 0, "PAST_DUE": 0, "ESCALATION": 0}

if GRAPH_AVAILABLE:
    print("[INFO] Microsoft Graph API configured — Outlook drafts will be created.")
else:
    print("[INFO] Graph API credentials not found — .txt drafts only (check .env file).")

for customer, group in actionable.groupby("Customer", sort=True):
    group = group.reset_index(drop=True)

    # Which invoices in this customer's group need a fresh email?
    new_invoices = [row for _, row in group.iterrows()
                    if not is_recently_contacted(row["Invoice"])]

    total_balance = group["Balance"].sum()
    invoice_count = len(group)
    # Max days past due across truly overdue invoices only (PRE_DUE = 0)
    max_days_past = int(group["Days_Past_Due"].clip(lower=0).max())
    email_type    = worst_category(group["Category"].tolist())

    # Always write to summary (skipped customers still appear)
    summary_rows.append({
        "Customer":          customer,
        "Total_Balance":     round(total_balance, 2),
        "Invoice_Count":     invoice_count,
        "Max_Days_Past_Due": max_days_past,
        "Email_Type":        email_type,
    })

    # Cooldown: skip if every invoice was recently contacted
    if not new_invoices:
        emails_skipped += 1
        continue

    # Build email text
    email_text = build_email(customer, group, total_balance, email_type)

    # Always write .txt fallback
    filepath = os.path.join(EMAIL_DIR, f"{safe_filename(customer)}_email.txt")
    with open(filepath, "w", encoding="utf-8") as fh:
        fh.write(email_text)

    # Optionally create Outlook draft (HTML, with bolding)
    _emails = group["Email"].fillna("").str.strip().loc[lambda s: s != ""] if "Email" in group.columns else []
    raw_addresses = _emails.iloc[0] if len(_emails) > 0 else ""
    html_subject, html_body = build_email_html(customer, group, total_balance, email_type)
    if create_outlook_draft(html_subject, html_body, raw_addresses):
        outlook_drafts += 1

    emails_written           += 1
    type_counts[email_type]  += 1

    # Log only the newly-contacted invoices
    followup = today + pd.Timedelta(days=COOLDOWN_DAYS)
    for row in new_invoices:
        new_log_rows.append({
            "Customer":        customer,
            "Invoice":         str(row["Invoice"]),
            "Email_Sent_Date": today.strftime("%m/%d/%Y"),
            "Followup_Date":   followup.strftime("%m/%d/%Y"),
            "Status":          "Email Sent",
            "Email_Type":      row["Category"],
            "Notes":           "",
        })


# ════════════════════════════════════════════════════════════════════════════
# 5.  PERSIST COLLECTIONS LOG
#     Append new rows; preserve existing Status / Notes on older entries.
# ════════════════════════════════════════════════════════════════════════════
if new_log_rows:
    if _db_available:
        _db.append_log_rows(new_log_rows)
    else:
        new_df = pd.DataFrame(new_log_rows, columns=LOG_COLUMNS)
        if os.path.exists(LOG_FILE):
            existing_log = pd.read_csv(LOG_FILE, dtype=str)
            for col in LOG_COLUMNS:
                if col not in existing_log.columns:
                    existing_log[col] = ""
            updated_log = pd.concat([existing_log, new_df], ignore_index=True)
        else:
            updated_log = new_df
        updated_log.to_csv(LOG_FILE, index=False)


# ════════════════════════════════════════════════════════════════════════════
# 6.  COLLECTIONS SUMMARY CSV
# ════════════════════════════════════════════════════════════════════════════
summary_df = (
    pd.DataFrame(summary_rows)
      .sort_values("Total_Balance", ascending=False)
      .reset_index(drop=True)
)
summary_df.to_csv(SUMMARY_FILE, index=False)


# ════════════════════════════════════════════════════════════════════════════
# 7.  EXECUTIVE AR SUMMARY REPORT
# ════════════════════════════════════════════════════════════════════════════
total_open_ar = df["Balance"].sum()
total_past_due_bal = past_due["Balance"].sum()
customers_pd  = past_due["Customer"].nunique()
invoices_pd   = len(past_due)

predue_inv    = (df["Category"] == "PRE_DUE").sum()
predue_cust   = df[df["Category"] == "PRE_DUE"]["Customer"].nunique()

top5     = summary_df.head(5)
sep      = "=" * 64
thin_sep = "-" * 64

report_lines = [
    sep,
    "  ACCOUNTS RECEIVABLE \u2014 EXECUTIVE SUMMARY",
    f"  Generated: {today.strftime('%B %d, %Y')}",
    sep,
    "",
    "  OVERVIEW",
    thin_sep,
    f"  {'Total Open AR':<32}  ${total_open_ar:>14,.2f}",
    f"  {'Total Past Due':<32}  ${total_past_due_bal:>14,.2f}",
    f"  {'Customers Past Due':<32}  {customers_pd:>15}",
    f"  {'Invoices Past Due':<32}  {invoices_pd:>15}",
    f"  {'Invoices Due Within 7 Days':<32}  {predue_inv:>15}",
    f"  {'Customers Due Within 7 Days':<32}  {predue_cust:>15}",
    "",
    "  TOP 5 CUSTOMERS BY ACTIONABLE BALANCE",
    thin_sep,
    f"  {'#':<4} {'Customer':<34} {'Balance':>12}  {'Type':<11} {'Max DPD':>7}",
    f"  {'-'*4} {'-'*34} {'-'*12}  {'-'*11} {'-'*7}",
]

for rank, (_, row) in enumerate(top5.iterrows(), start=1):
    name = textwrap.shorten(row["Customer"], width=34, placeholder="...")
    dpd  = int(row["Max_Days_Past_Due"])
    dpd_str = f"{dpd} days" if dpd > 0 else "due soon"
    report_lines.append(
        f"  {rank:<4} {name:<34} ${row['Total_Balance']:>11,.2f}"
        f"  {row['Email_Type']:<11} {dpd_str:>7}"
    )

report_lines += [
    "",
    thin_sep,
    f"  Collections log  : {LOG_FILE}",
    f"  Email drafts     : {EMAIL_DIR}/",
    f"  Summary CSV      : {SUMMARY_FILE}",
    sep,
    "",
]

report_text = "\n".join(report_lines)
with open(AR_REPORT, "w", encoding="utf-8") as fh:
    fh.write(report_text)


# ════════════════════════════════════════════════════════════════════════════
# 8.  COLLECTIONS PRIORITY REPORT
# ════════════════════════════════════════════════════════════════════════════
priority_df = (
    past_due
    .groupby("Customer", sort=False)
    .agg(
        Total_Balance    = ("Balance",      "sum"),
        Invoice_Count    = ("Invoice",      "count"),
        Max_Days_Past_Due= ("Days_Past_Due","max"),
    )
    .reset_index()
    .sort_values(["Total_Balance", "Max_Days_Past_Due"], ascending=[False, False])
    .reset_index(drop=True)
)

prio_lines = [
    sep,
    "  COLLECTIONS PRIORITY LIST",
    f"  Generated: {today.strftime('%B %d, %Y')}",
    sep,
    "  Customers ranked by total past-due balance, then max days past due.",
    thin_sep,
    f"  {'Rank':<5} {'Customer':<36} {'Balance':>12}  {'Inv':>3}  {'Max DPD':>9}",
    f"  {'-'*5} {'-'*36} {'-'*12}  {'-'*3}  {'-'*9}",
]

for rank, (_, row) in enumerate(priority_df.iterrows(), start=1):
    name = textwrap.shorten(row["Customer"], width=36, placeholder="...")
    prio_lines.append(
        f"  {rank:<5} {name:<36} ${row['Total_Balance']:>11,.2f}"
        f"  {int(row['Invoice_Count']):>3}  {int(row['Max_Days_Past_Due']):>6} days"
    )

prio_lines += [
    thin_sep,
    f"  Total customers past due : {len(priority_df)}",
    f"  Total past-due balance   : ${total_past_due_bal:,.2f}",
    sep,
    "",
]

priority_text = "\n".join(prio_lines)
with open(PRIORITY_REPORT, "w", encoding="utf-8") as fh:
    fh.write(priority_text)


# ════════════════════════════════════════════════════════════════════════════
# 9.  CONSOLE SUMMARY
# ════════════════════════════════════════════════════════════════════════════
print(report_text)

print("  EMAIL DISPATCH SUMMARY")
print(thin_sep)
print(f"  {'Payments detected (marked Paid)':<38}  {payments_detected:>4}")
print(f"  {'Pre-due reminders created':<38}  {type_counts['PRE_DUE']:>4}")
print(f"  {'Past-due notices created':<38}  {type_counts['PAST_DUE']:>4}")
print(f"  {'Escalation emails created':<38}  {type_counts['ESCALATION']:>4}")
print(f"  {'Customers skipped (cooldown)':<38}  {emails_skipped:>4}")
print(thin_sep)
print(f"  {'Total emails drafted':<38}  {emails_written:>4}")
print(f"  {'Outlook drafts created':<38}  {outlook_drafts:>4}")
print(f"  {'Total customers actioned':<38}  {len(summary_rows):>4}")
print(f"  {'New log entries added':<38}  {len(new_log_rows):>4}")
print(sep)
print()

# ── Top 10 priority accounts ─────────────────────────────────────────────────
print(sep)
print("  TOP 10 COLLECTION PRIORITIES")
print(thin_sep)
print(f"  {'Rank':<5} {'Customer':<36} {'Balance':>12}  {'Inv':>3}  {'Max DPD':>9}")
print(f"  {'-'*5} {'-'*36} {'-'*12}  {'-'*3}  {'-'*9}")
for rank, (_, row) in enumerate(priority_df.head(10).iterrows(), start=1):
    name = textwrap.shorten(row["Customer"], width=36, placeholder="...")
    print(
        f"  {rank:<5} {name:<36} ${row['Total_Balance']:>11,.2f}"
        f"  {int(row['Invoice_Count']):>3}  {int(row['Max_Days_Past_Due']):>6} days"
    )
print(thin_sep)
print(f"  Priority report saved : {PRIORITY_REPORT}")
print(sep)


# ════════════════════════════════════════════════════════════════════════════
# 10. WEEKLY COLLECTIONS UPDATE REPORT
# ════════════════════════════════════════════════════════════════════════════

# ── Rolling 7-day window ─────────────────────────────────────────────────────
week_start = today - pd.Timedelta(days=6)

# ── Create interactions file if missing (only needed for CSV fallback) ───────
if not _db_available and not os.path.exists(INTERACTIONS_FILE):
    pd.DataFrame(columns=["Date", "Customer", "Invoice", "Type", "Notes"]) \
      .to_csv(INTERACTIONS_FILE, index=False)

# ── Load interactions; degrade silently if empty or malformed ────────────────
# ia      — full DataFrame (used for customer-reply section)
# most_recent_contact — dict for Top 10 "Last Update" field
ia: pd.DataFrame = pd.DataFrame()
most_recent_contact: dict = {}
try:
    if _db_available:
        _ia_raw = _db.read_customer_interactions()
        ia = _ia_raw.fillna("") if _ia_raw is not None else pd.DataFrame()
    else:
        ia = pd.read_csv(INTERACTIONS_FILE, dtype=str).fillna("")
    if not ia.empty and {"Date", "Customer", "Notes"}.issubset(ia.columns):
        ia["_dt"] = pd.to_datetime(ia["Date"], errors="coerce")
        ia = ia.dropna(subset=["_dt"])
        for cust, grp in ia.groupby("Customer", sort=False):
            latest = grp.sort_values("_dt").iloc[-1]
            most_recent_contact[cust] = {
                "date":  latest["Date"],
                "notes": latest["Notes"] if latest["Notes"].strip() else "(no note)",
            }
except Exception:
    pass   # unreadable file → every customer shows "No recent contact logged"

# ── A. Customers contacted this week (from collections log) ──────────────────
week_type_counts = {"PRE_DUE": 0, "PAST_DUE": 0, "ESCALATION": 0}
if not log.empty:
    # Re-parse Email_Sent_Date (column may be datetime or string depending on path)
    _clog_sent_dt = pd.to_datetime(log["Email_Sent_Date"], format="%m/%d/%Y", errors="coerce")
    _week_sends = log[
        _clog_sent_dt.notna()
        & (_clog_sent_dt >= week_start)
        & (_clog_sent_dt <= today)
    ]
    for etype in week_type_counts:
        week_type_counts[etype] = int(
            (_week_sends["Email_Type"].fillna("") == etype).sum()
        )
week_total_contacted = sum(week_type_counts.values())

# ── B. Customer replies this week (from interactions file) ───────────────────
replies_this_week: list[dict] = []
if not ia.empty and "_dt" in ia.columns:
    _ia_week = (
        ia[ia["_dt"] >= week_start]
        .sort_values(["_dt", "Customer"])
    )
    for _, _row in _ia_week.iterrows():
        replies_this_week.append({
            "customer": _row.get("Customer", ""),
            "invoice":  _row.get("Invoice",  ""),
            "date":     _row["_dt"].strftime("%m/%d/%Y"),
            "notes":    _row.get("Notes", "") or "(no note)",
        })

# ── C. Payments logged this week (log entries where Status == "Paid") ────────
payments_this_week: list[dict] = []
if not log.empty:
    _plog_paid_dt = pd.to_datetime(log["Paid_Date"], format="%m/%d/%Y", errors="coerce")
    _paid_week = log[
        (log["Status"].fillna("") == "Paid")
        & _plog_paid_dt.notna()
        & (_plog_paid_dt >= week_start)
        & (_plog_paid_dt <= today)
        & log["Customer"].fillna("").str.strip().ne("")
        & log["Invoice"].fillna("").str.strip().ne("")
    ].copy()
    # Try to attach balance from current AR; fall back to None
    _inv_bal = df.set_index("Invoice")["Balance"] if not df.empty else pd.Series(dtype=float)
    for _, _pr in _paid_week.iterrows():
        _inv_str = str(_pr.get("Invoice", "") or "").strip()
        _bal = _inv_bal.get(_inv_str, None)
        _pd_dt = pd.to_datetime(_pr.get("Paid_Date", ""), format="%m/%d/%Y", errors="coerce")
        payments_this_week.append({
            "customer":  str(_pr.get("Customer", "") or ""),
            "invoice":   _inv_str,
            "paid_date": _pd_dt.strftime("%m/%d/%Y") if pd.notna(_pd_dt) else "",
            "balance":   _bal,
        })

# ── Build per-customer past-due summary (balance > 0, due date < today) ──────
wd_src = past_due[past_due["Balance"] > 0].copy()

weekly_rows = []
for customer, grp in wd_src.groupby("Customer", sort=False):
    grp = grp.sort_values("Days_Past_Due", ascending=False)
    weekly_rows.append({
        "customer":  customer,
        "total_bal": grp["Balance"].sum(),
        "max_dpd":   int(grp["Days_Past_Due"].max()),
        "invoices":  [
            {"num": str(r["Invoice"]),
             "bal": r["Balance"],
             "dpd": int(r["Days_Past_Due"])}
            for _, r in grp.iterrows()
        ],
    })

# Sort: Max_Days_Past_Due DESC, then Total_Balance DESC; keep top 10
weekly_rows.sort(key=lambda r: (-r["max_dpd"], -r["total_bal"]))
top10_weekly = weekly_rows[:10]

# ── Compose report ───────────────────────────────────────────────────────────
WIDE_SEP  = "=" * 64
ENTRY_SEP = "-" * 64
week_range = f"{week_start.strftime('%B %d, %Y')} \u2013 {today.strftime('%B %d, %Y')}"

lines = [
    WIDE_SEP,
    "  WEEKLY COLLECTIONS UPDATE",
    f"  Generated : {today.strftime('%B %d, %Y')}",
    f"  Week of   : {week_range}",
    WIDE_SEP,
    "",
    # ── Section 1: Customers Contacted ───────────────────────────────────
    "  CUSTOMERS CONTACTED THIS WEEK",
    ENTRY_SEP,
    f"  {'Total contacts':<28}  {week_total_contacted:>4}",
    f"  {'  Pre-due reminders':<28}  {week_type_counts['PRE_DUE']:>4}",
    f"  {'  Past-due notices':<28}  {week_type_counts['PAST_DUE']:>4}",
    f"  {'  Escalations':<28}  {week_type_counts['ESCALATION']:>4}",
    "",
    # ── Section 2: Customer Replies & Notes ──────────────────────────────
    "  CUSTOMER REPLIES & NOTES",
    ENTRY_SEP,
]

if replies_this_week:
    lines.append(
        f"  {'Date':<12} {'Invoice':<10} {'Customer':<34} Notes"
    )
    lines.append(
        f"  {'-'*12} {'-'*10} {'-'*34} {'-'*30}"
    )
    for r in replies_this_week:
        cust_short = textwrap.shorten(r["customer"], width=34, placeholder="...")
        # Wrap long notes at ~60 chars, indent continuation lines
        note_lines = textwrap.wrap(r["notes"], width=60) or ["(no note)"]
        lines.append(
            f"  {r['date']:<12} {r['invoice']:<10} {cust_short:<34} {note_lines[0]}"
        )
        indent = " " * (2 + 12 + 1 + 10 + 1 + 34 + 1)
        for nl in note_lines[1:]:
            lines.append(f"{indent}{nl}")
else:
    lines.append("  (none logged this week)")

lines += [
    "",
    # ── Section 3: Payments Logged This Week ─────────────────────────────
    "  PAYMENTS LOGGED THIS WEEK",
    ENTRY_SEP,
]

if payments_this_week:
    lines.append(
        f"  {'Paid Date':<12} {'Invoice':<10} {'Customer':<34} {'Amount':>12}"
    )
    lines.append(
        f"  {'-'*12} {'-'*10} {'-'*34} {'-'*12}"
    )
    for p in payments_this_week:
        cust_short = textwrap.shorten(p["customer"], width=34, placeholder="...")
        amt = f"${p['balance']:>10,.2f}" if p["balance"] is not None else f"{'—':>11}"
        lines.append(
            f"  {p['paid_date']:<12} {p['invoice']:<10} {cust_short:<34} {amt}"
        )
else:
    lines.append("  (none detected this week)")

lines += [
    "",
    # ── Section 4: Total Past Due ─────────────────────────────────────────
    "  TOTAL PAST DUE",
    ENTRY_SEP,
    f"  {'Past-due balance':<28}  ${total_past_due_bal:>14,.2f}",
    f"  {'Customers past due':<28}  {customers_pd:>15}",
    "",
    WIDE_SEP,
    "",
    # ── Section 5: Top 10 Past-Due Accounts ──────────────────────────────
    "  TOP 10 PAST-DUE ACCOUNTS  (ranked by days overdue)",
    WIDE_SEP,
    "",
]

for rank, item in enumerate(top10_weekly, start=1):
    customer = item["customer"]
    contact  = most_recent_contact.get(customer)
    last_update = (
        f"{contact['date']} \u2013 {contact['notes']}"
        if contact else "No recent contact logged"
    )

    lines.append(
        f"{rank}. {customer} \u2014 ${item['total_bal']:,.2f}"
        f" \u2014 {item['max_dpd']} days past due"
    )
    lines.append("")
    lines.append("Invoices")
    for inv in item["invoices"]:
        lines.append(f"{inv['num']} \u2013 ${inv['bal']:,.2f} \u2013 {inv['dpd']} days")
    lines.append("")
    lines.append("Last Update")
    lines.append(last_update)
    lines.append("")
    lines.append(ENTRY_SEP)
    lines.append("")

weekly_text = "\n".join(lines)
with open(WEEKLY_REPORT, "w", encoding="utf-8") as fh:
    fh.write(weekly_text)

print(f"[INFO] Weekly collections report generated:")
print(f"       {WEEKLY_REPORT}")
