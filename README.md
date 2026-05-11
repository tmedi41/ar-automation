# AR Collections Automation Platform

A production-deployed accounts receivable collections automation platform built for a small industrial services company in Houston, TX.

---

## The Problem

Manual AR collections meant hours of weekly outreach — reviewing aging reports, drafting individual emails, tracking responses, and compiling weekly reports for leadership. This platform automates that entire workflow.

---

## What It Does

- Ingests AR export data and identifies past due accounts
- Auto-generates personalized collection email drafts via Microsoft Graph API (Outlook integration)
- Uses Anthropic AI API to summarize customer replies and flag key information
- Real-time web dashboard gives leadership visibility into receivables, aging buckets, payment status, and contact history
- Automated weekly AR report generation

---

## Impact

- Eliminated approximately **15 hours per week** of manual outreach
- Supporting collection of **$257K+** in monthly receivables
- Tracks **30+ active accounts** across priority tiers

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python, Flask |
| Database | PostgreSQL |
| Integrations | Microsoft Graph API, Anthropic AI API |
| Frontend | Custom fintech-inspired web dashboard |
| Deployment | Railway |

---

## Dashboard Views

- **Overview** — Open AR, past due accounts, weekly contacts
- **Customer Replies** — AI-summarized inbound responses
- **Weekly Report** — Auto-generated leadership summary
- **Payment History** — Collections trends and top overdue accounts

---

## Setup

1. Clone the repo
2. Copy `.env.example` to `.env` and fill in your credentials
3. Run `pip install -r requirements.txt`
4. Run `python app.py`
