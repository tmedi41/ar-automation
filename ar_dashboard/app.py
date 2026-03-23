"""
ar_dashboard/app.py
--------------------
Flask web dashboard for the AR Automation system.
Reads reports/, database/, and data/ from the parent AR_Automation directory.
Run:  python3 ar_dashboard/app.py
"""

import os
import re
import subprocess
from datetime import datetime, timedelta

import msal
import requests as http_requests
import anthropic
import pandas as pd
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

load_dotenv(os.path.join(BASE_DIR, ".env"), override=True)

app = Flask(__name__)


# ── Data helpers ─────────────────────────────────────────────────────────────

def _read_csv(rel_path: str) -> pd.DataFrame:
    path = os.path.join(BASE_DIR, rel_path)
    if not os.path.exists(path):
        return pd.DataFrame()
    return pd.read_csv(path, dtype=str)


def get_metrics() -> dict:
    metrics = {
        "total_open_ar":     "$0.00",
        "total_past_due":    "$0.00",
        "payments_this_week": 0,
        "drafts_this_week":   0,
        "generated":          "",
        "customers_past_due": 0,
    }

    # ── ar_summary.txt ───────────────────────────────────────────────────────
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

    # ── collections_log.csv ──────────────────────────────────────────────────
    log = _read_csv("database/collections_log.csv")
    if not log.empty:
        today      = datetime.today().date()
        week_start = today - timedelta(days=6)

        log["_paid_dt"] = pd.to_datetime(log.get("Paid_Date", ""), format="%m/%d/%Y", errors="coerce")
        log["_sent_dt"] = pd.to_datetime(log.get("Email_Sent_Date", ""), format="%m/%d/%Y", errors="coerce")

        has_customer = log["Customer"].fillna("").str.strip().ne("")
        has_invoice  = log["Invoice"].fillna("").str.strip().ne("")

        paid_mask = (
            (log["Status"].fillna("") == "Paid")
            & log["_paid_dt"].notna()
            & (log["_paid_dt"].dt.date >= week_start)
            & (log["_paid_dt"].dt.date <= today)
            & has_customer & has_invoice
        )
        metrics["payments_this_week"] = int(paid_mask.sum())

        sent_mask = (
            log["_sent_dt"].notna()
            & (log["_sent_dt"].dt.date >= week_start)
            & (log["_sent_dt"].dt.date <= today)
            & has_customer
        )
        metrics["drafts_this_week"] = int(sent_mask.sum())

    return metrics


def get_priority_customers() -> list[dict]:
    df = _read_csv("reports/collections_summary.csv")
    if df.empty:
        return []
    df["Total_Balance"]     = pd.to_numeric(df["Total_Balance"],     errors="coerce").fillna(0)
    df["Max_Days_Past_Due"] = pd.to_numeric(df["Max_Days_Past_Due"], errors="coerce").fillna(0).astype(int)
    df["Invoice_Count"]     = pd.to_numeric(df["Invoice_Count"],     errors="coerce").fillna(0).astype(int)
    df = df.sort_values("Total_Balance", ascending=False).head(20)

    rows = []
    for _, row in df.iterrows():
        rows.append({
            "customer": row["Customer"],
            "balance":  f"${row['Total_Balance']:,.2f}",
            "balance_raw": float(row["Total_Balance"]),
            "invoices": int(row["Invoice_Count"]),
            "max_dpd":  int(row["Max_Days_Past_Due"]),
            "type":     str(row.get("Email_Type", "")),
        })
    return rows


def get_recent_replies() -> list[dict]:
    df = _read_csv("data/customer_interactions.csv")
    if df.empty:
        return []
    df = df.fillna("").copy()
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
    path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    if not os.path.exists(path):
        return []
    df = pd.read_csv(path, dtype=str).fillna("")
    df = df[df["Notes"].str.strip() != ""]
    if df.empty:
        return []
    df["_sort"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.sort_values("_sort", ascending=False).drop(columns=["_sort"])
    return df[["Date", "Invoice", "Customer", "Notes"]].rename(
        columns={"Date": "date", "Invoice": "invoice", "Customer": "customer", "Notes": "notes"}
    ).to_dict(orient="records")


def get_weekly_snapshot() -> str:
    path = os.path.join(BASE_DIR, "reports", "weekly_collections_update.txt")
    if not os.path.exists(path):
        return "(no weekly report found)"
    return open(path, encoding="utf-8").read()


def parse_weekly_report() -> dict:
    path = os.path.join(BASE_DIR, "reports", "weekly_collections_update.txt")
    result = {
        "generated": "", "week_of": "",
        "total_contacts": 0, "pre_due": 0, "past_due_notices": 0, "escalations": 0,
        "replies": [], "past_due_accounts": [],
    }
    if not os.path.exists(path):
        return result

    lines = open(path, encoding="utf-8").read().splitlines()

    # Header
    for line in lines:
        m = re.search(r"Generated\s*:\s*(.+)", line)
        if m:
            result["generated"] = m.group(1).strip()
        m = re.search(r"Week of\s*:\s*(.+)", line)
        if m:
            result["week_of"] = m.group(1).strip()

    # Contact summary
    for line in lines:
        m = re.search(r"Total contacts\s+(\d+)", line)
        if m:
            result["total_contacts"] = int(m.group(1))
        m = re.search(r"Pre-due reminders\s+(\d+)", line)
        if m:
            result["pre_due"] = int(m.group(1))
        m = re.search(r"Past-due notices\s+(\d+)", line)
        if m:
            result["past_due_notices"] = int(m.group(1))
        m = re.search(r"Escalations\s+(\d+)", line)
        if m:
            result["escalations"] = int(m.group(1))

    # Customer replies & notes (fixed-width columns: date[2:14], invoice[15:25], customer[26:60], notes[61:])
    in_replies = False
    current_reply = None
    for line in lines:
        if "CUSTOMER REPLIES & NOTES" in line:
            in_replies = True
            continue
        if in_replies and "PAYMENTS LOGGED" in line:
            if current_reply:
                result["replies"].append(current_reply)
            break
        if not in_replies:
            continue
        if re.match(r'\s*[-=]+\s*$', line) or re.match(r'\s*Date\s+Invoice', line):
            continue
        if not line.strip():
            continue

        if re.match(r'  \d{2}/\d{2}/\d{4}', line):
            if current_reply:
                result["replies"].append(current_reply)
            current_reply = {
                "date":     line[2:14].strip(),
                "invoice":  line[15:25].strip(),
                "customer": line[26:60].strip(),
                "notes":    line[61:].strip() if len(line) > 61 else "",
            }
        elif current_reply and len(line) > 61 and not line[2:61].strip():
            current_reply["notes"] += " " + line[61:].strip()

    # Past due accounts
    in_accounts = False
    current_account = None
    account_phase = None
    dash_re = re.compile(r'[—–-]')

    for line in lines:
        if "TOP 10 PAST-DUE ACCOUNTS" in line:
            in_accounts = True
            continue
        if not in_accounts:
            continue

        stripped = line.strip()
        m = re.match(r'(\d+)\.\s+(.+?)\s+[—–]\s+\$([0-9,]+\.?\d*)\s+[—–]\s+(\d+)\s+days past due', stripped)
        if m:
            if current_account:
                result["past_due_accounts"].append(current_account)
            current_account = {
                "rank": int(m.group(1)),
                "customer": m.group(2).strip(),
                "total": f"${float(m.group(3).replace(',', '')):,.2f}",
                "days_past_due": int(m.group(4)),
                "invoices": [],
                "last_update": "",
            }
            account_phase = None
            continue

        if current_account is None:
            continue
        if stripped == "Invoices":
            account_phase = "invoices"
            continue
        if stripped == "Last Update":
            account_phase = "last_update"
            continue
        if re.match(r'^[-=]+$', stripped) or not stripped:
            continue

        if account_phase == "invoices":
            m = re.match(r'(\d+)\s+[—–]\s+\$([0-9,]+\.?\d*)\s+[—–]\s+(\d+)\s+days', stripped)
            if m:
                current_account["invoices"].append({
                    "number": m.group(1),
                    "amount": f"${float(m.group(2).replace(',', '')):,.2f}",
                    "days": int(m.group(3)),
                })
        elif account_phase == "last_update":
            current_account["last_update"] = stripped

    if current_account:
        result["past_due_accounts"].append(current_account)

    return result


def enrich_last_updates(past_due_accounts: list) -> list:
    """Override last_update on each past-due account using the most recent
    matching entry in customer_interactions.csv, keyed by invoice number."""
    path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    if not os.path.exists(path):
        return past_due_accounts

    df = pd.read_csv(path, dtype=str).fillna("")
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


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    weekly_data = parse_weekly_report()
    weekly_data["past_due_accounts"] = enrich_last_updates(weekly_data["past_due_accounts"])
    return render_template(
        "index.html",
        metrics=get_metrics(),
        priority=get_priority_customers(),
        replies=get_recent_replies(),
        weekly=get_weekly_snapshot(),
        weekly_data=weekly_data,
        customer_replies=get_customer_replies(),
    )


@app.route("/run-automation", methods=["POST"])
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
def upload_ar():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided."})
    f = request.files["file"]
    if not f.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "File must be a .csv"})
    dest = os.path.join(BASE_DIR, "exports", "ar_aging.csv")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    f.save(dest)
    return jsonify({"ok": True})


@app.route("/summarize-reply", methods=["POST"])
def summarize_reply():
    data    = request.get_json()
    invoice = (data.get("invoice") or "").strip()
    reply   = (data.get("reply")   or "").strip()
    if not invoice or not reply:
        return jsonify({"ok": False, "error": "Invoice and reply are required."})

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return jsonify({"ok": False, "error": "ANTHROPIC_API_KEY not set in .env"})

    client  = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=60,
        system=(
            "Summarize customer replies in one phrase, 15 words or fewer. "
            "Capture only the key action or commitment — e.g. 'Will pay via ACH on April 6' or "
            "'Payment terms are 60 days, due April 7' or 'Invoice in system to be paid.' "
            "No subjects, no pleasantries, no signatures, no filler. Just the key fact."
        ),
        messages=[{
            "role": "user",
            "content": reply,
        }],
    )
    summary = message.content[0].text.strip()
    return jsonify({"ok": True, "summary": summary})


@app.route("/log-reply", methods=["POST"])
def log_reply():
    data     = request.get_json()
    invoice  = (data.get("invoice")  or "").strip()
    customer = (data.get("customer") or "").strip()
    summary  = (data.get("summary")  or "").strip()
    if not invoice or not summary:
        return jsonify({"ok": False, "error": "Invoice and summary are required."})

    csv_path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    today    = datetime.today().strftime("%Y-%m-%d")
    new_row  = pd.DataFrame([{
        "Date":     today,
        "Customer": customer,
        "Invoice":  invoice,
        "Type":     "Customer Reply",
        "Notes":    summary,
    }])

    if os.path.exists(csv_path):
        existing = pd.read_csv(csv_path, dtype=str)
        updated  = pd.concat([existing, new_row], ignore_index=True)
    else:
        updated = new_row

    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    updated.to_csv(csv_path, index=False)
    return jsonify({"ok": True})


@app.route("/delete-reply", methods=["POST"])
def delete_reply():
    data     = request.get_json()
    raw_date = (data.get("raw_date") or "").strip()
    invoice  = (data.get("invoice")  or "").strip()
    customer = (data.get("customer") or "").strip()
    notes    = (data.get("notes")    or "").strip()

    csv_path = os.path.join(BASE_DIR, "data", "customer_interactions.csv")
    if not os.path.exists(csv_path):
        return jsonify({"ok": False, "error": "No interactions file found."})

    df = pd.read_csv(csv_path, dtype=str).fillna("")
    mask = (
        (df["Date"].str.strip()     == raw_date) &
        (df["Invoice"].str.strip()  == invoice)  &
        (df["Customer"].str.strip() == customer) &
        (df["Notes"].str.strip()    == notes)
    )
    idx = df.index[mask]
    if idx.empty:
        return jsonify({"ok": False, "error": "Entry not found."})

    df = df.drop(idx[0])
    df.to_csv(csv_path, index=False)
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port)
