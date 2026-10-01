#!/usr/bin/env python3
"""
Cold Email Automator — v2 (Dynamic, Verified, Anti-Hallucination)
=================================================================
- Dynamically detects ALL populated rows in Excel (never hardcodes counts)
- Per-recipient JSON status tracking (not sequential history.txt)
- IMAP verification against Gmail Sent folder after every send
- Full audit report with reconciled counts
- Self-correcting: re-runs always pick up new rows + previously failed/missing rows
"""

import pandas as pd
import smtplib
import imaplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import random
import os
import sys
import io
from datetime import datetime, timezone

# ── UTF-8 fix for Windows console ──────────────────────────────────────────
if sys.stdout and sys.stdout.buffer:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr and sys.stderr.buffer:
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import config

# ── File paths ──────────────────────────────────────────────────────────────
STATUS_DB_FILE  = "status_db.json"
SKIPPED_LOG     = "skipped_log.txt"
AUDIT_LOG       = "audit_report.txt"
HISTORY_FILE    = "history.txt"

# ── Status constants ────────────────────────────────────────────────────────
SENT_VERIFIED        = "SENT_VERIFIED"
NOT_SENT             = "NOT_SENT"
VERIFICATION_REQUIRED = "VERIFICATION_REQUIRED"
SKIPPED              = "SKIPPED"


# ════════════════════════════════════════════════════════════════════════════
# STATUS DATABASE (per-recipient JSON)
# ════════════════════════════════════════════════════════════════════════════

def load_status_db() -> dict:
    """Load per-recipient status database from disk."""
    if os.path.exists(STATUS_DB_FILE):
        with open(STATUS_DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_status_db(db: dict):
    """Persist status database to disk immediately after every change."""
    with open(STATUS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)

def log_skipped(sno, name, email, reason):
    """Append a skipped-contact record to skipped_log.txt."""
    with open(SKIPPED_LOG, "a", encoding="utf-8") as f:
        f.write(f"[{_now()}] SNo {sno} | {name} | {email} | {reason}\n")


# ════════════════════════════════════════════════════════════════════════════
# EXCEL — ALWAYS DYNAMIC
# ════════════════════════════════════════════════════════════════════════════

def load_excel() -> pd.DataFrame:
    """
    Dynamically load the Excel file.
    - Never hardcodes row count or range.
    - Handles merged-header sheets (detects real column row automatically).
    - Drops completely empty rows.
    - Auto-generates SNo if the column is absent.
    """
    try:
        df = pd.read_excel(config.EXCEL_PATH)

        # Detect real column headers if they're on row 1 instead of row 0
        real_cols = {"Name", "Personal Email", "Email"}
        if not real_cols.intersection(set(df.columns)):
            if not df.empty and real_cols.intersection(set(df.iloc[0].values)):
                df = pd.read_excel(config.EXCEL_PATH, header=1)

        # Drop rows that are entirely NaN
        df = df.dropna(how="all").reset_index(drop=True)

        # Auto-generate SNo if missing
        if "SNo" not in df.columns:
            df["SNo"] = range(1, len(df) + 1)

        df["SNo"] = df["SNo"].astype(int)
        return df

    except Exception as exc:
        print(f"❌ Cannot load Excel file '{config.EXCEL_PATH}': {exc}")
        sys.exit(1)


# ════════════════════════════════════════════════════════════════════════════
# MIGRATION — history.txt → status_db.json (first-run only)
# ════════════════════════════════════════════════════════════════════════════

def migrate_from_history_if_needed(df: pd.DataFrame, suppressed_emails: set, suppressed_names: set):
    """
    One-time migration: if status_db.json is absent but history.txt exists,
    build the DB from the old sequential tracker.  Rows up to last_sent SNo
    that have no email are marked SKIPPED; suppressed rows are SKIPPED;
    everything else that was 'sent' is marked VERIFICATION_REQUIRED so IMAP
    can later confirm them.
    """
    if os.path.exists(STATUS_DB_FILE):
        return  # Already migrated

    OLD_HISTORY = "history.txt"
    if not os.path.exists(OLD_HISTORY):
        return  # Fresh install — nothing to migrate

    with open(OLD_HISTORY, "r") as f:
        content = f.read().strip()
    last_sent_sno = int(content) if content.isdigit() else 0

    if last_sent_sno == 0:
        return

    print(f"🔄 Migrating from history.txt (last sent SNo = {last_sent_sno}) → status_db.json …")
    db = {}

    for _, row in df.iterrows():
        sno = str(int(row["SNo"]))
        int_sno = int(sno)
        if int_sno > last_sent_sno:
            break  # These were never sent

        hr_name = _get_name(row)
        to_email = _get_email(row)
        company  = _get_company(row)

        # Suppression check
        if to_email and (to_email.lower() in suppressed_emails or hr_name.lower() in suppressed_names):
            db[sno] = _record(sno, hr_name, to_email or "N/A", company, SKIPPED, skip_reason="On suppression list (migrated)")
            continue

        # Missing email
        if not to_email:
            db[sno] = _record(sno, hr_name, "N/A", company, SKIPPED, skip_reason="No email in Excel (migrated)")
            continue

        # Missing company
        if not company:
            db[sno] = _record(sno, hr_name, to_email, "N/A", SKIPPED, skip_reason="No company in Excel (migrated)")
            continue

        # Sent — pending IMAP verification
        db[sno] = _record(sno, hr_name, to_email, company, VERIFICATION_REQUIRED)

    save_status_db(db)
    print(f"   ✅ Migration complete. {len(db)} rows pre-loaded into status_db.json.")


# ════════════════════════════════════════════════════════════════════════════
# SUPPRESSION LIST
# ════════════════════════════════════════════════════════════════════════════

def load_suppression_list():
    """
    Parse DO_NOT_EMAIL.md to build suppressed email + name sets.
    Returns: (suppressed_emails: set[str], suppressed_names: set[str])
    """
    suppressed_emails, suppressed_names = set(), set()
    path = getattr(config, "DO_NOT_EMAIL_PATH", "DO_NOT_EMAIL.md")
    if not os.path.exists(path):
        return suppressed_emails, suppressed_names

    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("|") and not line.startswith("|---") and not line.startswith("| Name"):
                    parts = [p.strip() for p in line.split("|")[1:-1]]
                    if len(parts) >= 2:
                        name_v, email_v = parts[0], parts[1]
                        if name_v and name_v.lower() != "name":
                            suppressed_names.add(name_v.strip().lower())
                        for em in email_v.replace(";", ",").split(","):
                            em = em.strip().lower()
                            if "@" in em:
                                suppressed_emails.add(em)
                else:
                    for word in line.split():
                        w = word.strip("|,;()[]`'\"").lower()
                        if "@" in w and "." in w:
                            suppressed_emails.add(w)
    except Exception as exc:
        print(f"⚠️  Could not read suppression list: {exc}")

    return suppressed_emails, suppressed_names


# ════════════════════════════════════════════════════════════════════════════
# EMAIL SENDING
# ════════════════════════════════════════════════════════════════════════════

def create_email_body(hr_name: str, company_name: str) -> str:
    return config.EMAIL_TEMPLATE.format(
        HR_NAME=hr_name,
        COMPANY_NAME=company_name,
        YOUR_NAME=config.YOUR_NAME,
        YOUR_COURSE=config.YOUR_COURSE,
        YOUR_COLLEGE=config.YOUR_COLLEGE,
        YOUR_LINKEDIN=config.YOUR_LINKEDIN,
        YOUR_GITHUB=config.YOUR_GITHUB,
        YOUR_PORTFOLIO=config.YOUR_PORTFOLIO,
        YOUR_PHONE=config.YOUR_PHONE,
        YOUR_EMAIL=config.YOUR_EMAIL,
    )

def send_email(to_email: str, subject: str, body: str):
    """Send email via SMTP SSL. Raises on failure."""
    msg = MIMEMultipart()
    msg["From"]    = config.YOUR_EMAIL
    msg["To"]      = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if os.path.exists(config.RESUME_PATH):
        with open(config.RESUME_PATH, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(config.RESUME_PATH))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(config.RESUME_PATH)}"'
            msg.attach(part)
    else:
        print(f"⚠️  Resume not found at '{config.RESUME_PATH}'. Sending without attachment.")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(config.YOUR_EMAIL, config.APP_PASSWORD)
        server.send_message(msg)


# ════════════════════════════════════════════════════════════════════════════
# IMAP VERIFICATION — checks Gmail Sent folder
# ════════════════════════════════════════════════════════════════════════════

def verify_via_imap(email_addresses: list[str]) -> dict[str, bool]:
    """
    Check Gmail's '[Gmail]/Sent Mail' folder for each address.
    Returns {email: True} only when a matching sent message is ACTUALLY found.
    Returns {email: False} on any error or when not found.

    ANTI-HALLUCINATION: This function NEVER fabricates results.
    Only True when IMAP search returns message UIDs.
    """
    results = {addr: False for addr in email_addresses}

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(config.YOUR_EMAIL, config.APP_PASSWORD)
        status, _ = mail.select('"[Gmail]/Sent Mail"')

        if status != "OK":
            print("⚠️  IMAP: Could not select Sent Mail folder.")
            mail.logout()
            return results

        for addr in email_addresses:
            try:
                # Search by recipient TO field
                typ, data = mail.search(None, f'TO "{addr}"')
                if typ == "OK" and data and data[0]:
                    results[addr] = True   # Evidence found in Sent folder
            except Exception as search_err:
                print(f"   ⚠️  IMAP search failed for {addr}: {search_err}")

        mail.logout()

    except imaplib.IMAP4.error as imap_err:
        print(f"⚠️  IMAP connection/login failed: {imap_err}")
        print("   All sent emails will be marked VERIFICATION_REQUIRED.")
    except Exception as exc:
        print(f"⚠️  Unexpected IMAP error: {exc}")

    return results


# ════════════════════════════════════════════════════════════════════════════
# HUMAN-READABLE HISTORY FILE
# ════════════════════════════════════════════════════════════════════════════

def write_history_txt(status_db: dict):
    """
    Writes a human-readable history.txt after every run.
    Sorted by SNo — covers every recipient so you can manually verify.

    Columns: SNo | Name | Email | Company | Status | Sent At | Verified At
    """
    STATUS_EMOJI = {
        SENT_VERIFIED:         "✅ SENT_VERIFIED",
        VERIFICATION_REQUIRED: "⚠️  VERIFICATION_REQUIRED",
        NOT_SENT:              "❌ NOT_SENT",
        SKIPPED:               "⏭️  SKIPPED",
    }

    # Sort records by SNo numerically
    sorted_records = sorted(status_db.items(), key=lambda x: int(x[0]))

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines   = []
    lines.append(f"Cold Email Automator — History Log")
    lines.append(f"Last updated : {now_str}")
    lines.append(f"Total records: {len(sorted_records)}")
    lines.append("")
    lines.append(
        f"{'SNo':<5} "
        f"{'Name':<30} "
        f"{'Email':<42} "
        f"{'Company':<28} "
        f"{'Status':<24} "
        f"{'Sent At (UTC)':<26} "
        f"{'Verified At (UTC)'}"
    )
    lines.append("-" * 170)

    for sno, rec in sorted_records:
        name      = str(rec.get("name",    "")).strip()[:29]
        email     = str(rec.get("email",   "")).strip()[:41]
        company   = str(rec.get("company", "")).strip()[:27]
        status    = STATUS_EMOJI.get(rec.get("status", NOT_SENT), rec.get("status", ""))
        sent_at   = str(rec.get("sent_at",     ""))[:25]
        verf_at   = str(rec.get("verified_at", ""))[:25]
        skip_note = rec.get("skip_reason", "")

        row_line = (
            f"{sno:<5} "
            f"{name:<30} "
            f"{email:<42} "
            f"{company:<28} "
            f"{status:<24} "
            f"{sent_at:<26} "
            f"{verf_at}"
        )
        if skip_note:
            row_line += f"  [{skip_note}]"
        lines.append(row_line)

    lines.append("")
    lines.append("-" * 170)

    # Status summary
    from collections import Counter
    status_counts = Counter(r.get("status", NOT_SENT) for _, r in sorted_records)
    lines.append(f"SUMMARY  |  "
                 f"SENT_VERIFIED: {status_counts[SENT_VERIFIED]}  |  "
                 f"VERIFICATION_REQUIRED: {status_counts[VERIFICATION_REQUIRED]}  |  "
                 f"NOT_SENT: {status_counts[NOT_SENT]}  |  "
                 f"SKIPPED: {status_counts[SKIPPED]}")

    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"📄 history.txt updated — {len(sorted_records)} records written.")


# ════════════════════════════════════════════════════════════════════════════
# FIELD EXTRACTORS
# ════════════════════════════════════════════════════════════════════════════

def _get_name(row) -> str:
    v = row.get("Name")
    return str(v).strip() if pd.notna(v) and str(v).strip() not in ("", "nan") else "Hiring Manager"

def _get_email(row) -> str | None:
    """Returns first valid email from Personal Email or Email column, or None."""
    for col in ("Personal Email", "Email"):
        v = row.get(col)
        if pd.notna(v) and str(v).strip() not in ("", "nan"):
            first = str(v).split(";")[0].strip()
            if "@" in first:
                return first
    return None

def _get_company(row) -> str | None:
    """Returns company name or None if absent."""
    for col in ("Company Name", "Company"):
        v = row.get(col)
        if pd.notna(v) and str(v).strip() not in ("", "nan"):
            return str(v).strip()
    return None

def _record(sno, name, email, company, status, **extra) -> dict:
    return {"sno": sno, "name": name, "email": email, "company": company,
            "status": status, "updated_at": _now(), **extra}

def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    if not config.SENDING_ENABLED:
        print("⛔ Email sending is disabled in config.py. Set SENDING_ENABLED = True to proceed.")
        return

    if not config.YOUR_EMAIL or not config.APP_PASSWORD:
        print("❌ YOUR_EMAIL or APP_PASSWORD is missing in .env")
        sys.exit(1)

    # ── 1. Load Excel dynamically ──────────────────────────────────────────
    df = load_excel()
    total_detected = len(df)
    print(f"\n📂 Excel loaded: {total_detected} recipients detected (dynamic)")

    # ── 2. Load suppression list ───────────────────────────────────────────
    suppressed_emails, suppressed_names = load_suppression_list()

    # ── 3. One-time migration from old history.txt ─────────────────────────
    migrate_from_history_if_needed(df, suppressed_emails, suppressed_names)

    # ── 4. Load status DB ─────────────────────────────────────────────────
    status_db = load_status_db()

    # ── 5. Categorise every row ────────────────────────────────────────────
    to_send    = []   # rows needing email
    already_ok = []   # rows already SENT_VERIFIED
    need_verify = []  # rows VERIFICATION_REQUIRED (sent but unverified)

    for _, row in df.iterrows():
        sno  = str(int(row["SNo"]))
        rec  = status_db.get(sno, {})
        curr = rec.get("status", NOT_SENT)

        hr_name  = _get_name(row)
        to_email = _get_email(row)
        company  = _get_company(row)

        # Already confirmed — skip
        if curr == SENT_VERIFIED:
            already_ok.append(sno)
            continue

        # Already marked SKIPPED in DB — respect it unless new email found
        if curr == SKIPPED and not to_email:
            continue

        # ── Validate fields ────────────────────────────────────────────────
        if not company:
            reason = "Company name missing — skipped to avoid 'your company' placeholder"
            if curr != SKIPPED:
                print(f"⚠️  SNo {sno} ({hr_name}): {reason}")
                log_skipped(sno, hr_name, to_email or "N/A", reason)
                status_db[sno] = _record(sno, hr_name, to_email or "N/A", "N/A", SKIPPED, skip_reason=reason)
                save_status_db(status_db)
            continue

        if not to_email:
            reason = "No email address in Excel"
            if curr != SKIPPED:
                print(f"⚠️  SNo {sno} ({hr_name} @ {company}): {reason}")
                log_skipped(sno, hr_name, "N/A", reason)
                status_db[sno] = _record(sno, hr_name, "N/A", company, SKIPPED, skip_reason=reason)
                save_status_db(status_db)
            continue

        # ── Suppression check ──────────────────────────────────────────────
        if to_email.lower() in suppressed_emails or hr_name.lower() in suppressed_names:
            reason = f"On suppression list ({getattr(config, 'DO_NOT_EMAIL_PATH', 'DO_NOT_EMAIL.md')})"
            if curr != SKIPPED:
                print(f"⏭️  SNo {sno} ({hr_name}): {reason}")
                log_skipped(sno, hr_name, to_email, reason)
                status_db[sno] = _record(sno, hr_name, to_email, company, SKIPPED, skip_reason=reason)
                save_status_db(status_db)
            continue

        # ── VERIFICATION_REQUIRED: batch for re-verification ───────────────
        if curr == VERIFICATION_REQUIRED:
            need_verify.append({"sno": sno, "email": to_email})
            continue

        # ── NOT_SENT: needs email ──────────────────────────────────────────
        to_send.append({"sno": sno, "hr_name": hr_name,
                        "company_name": company, "to_email": to_email})

    # ── 6. Re-verify VERIFICATION_REQUIRED rows first ─────────────────────
    if need_verify:
        print(f"\n🔍 Re-verifying {len(need_verify)} previously-sent emails via IMAP …")
        emails_to_check = [x["email"] for x in need_verify]
        imap_results = verify_via_imap(emails_to_check)

        for item in need_verify:
            sno   = item["sno"]
            email = item["email"]
            rec   = status_db.get(sno, {})

            if imap_results.get(email):
                status_db[sno] = {**rec, "status": SENT_VERIFIED,
                                  "verified_at": _now(), "updated_at": _now()}
                already_ok.append(sno)
                print(f"   ✅ SENT_VERIFIED: SNo {sno} — {email}")
            else:
                print(f"   ⚠️  VERIFICATION_REQUIRED: SNo {sno} — {email} (not found in Sent folder)")

        save_status_db(status_db)

    # ── 7. Send emails ────────────────────────────────────────────────────
    if not to_send:
        print("\n✅ Nothing new to send. All pending recipients already processed.")
    else:
        batch = to_send[:config.BATCH_SIZE]
        remaining_after = len(to_send) - len(batch)
        print(f"\n🚀 Sending batch: {len(batch)} emails"
              f"{f' ({remaining_after} more queued for next run)' if remaining_after else ''} …\n")

        sent_this_run = []

        for i, item in enumerate(batch):
            sno, hr_name = item["sno"], item["hr_name"]
            company_name, to_email = item["company_name"], item["to_email"]

            body = create_email_body(hr_name, company_name)
            print(f"[{sno}] Sending to {hr_name} ({to_email}) @ {company_name} …")

            try:
                send_email(to_email, config.SUBJECT, body)
                status_db[sno] = _record(sno, hr_name, to_email, company_name,
                                         VERIFICATION_REQUIRED, sent_at=_now())
                save_status_db(status_db)
                sent_this_run.append({"sno": sno, "email": to_email})

                if i < len(batch) - 1:
                    wait = random.randint(config.MIN_WAIT, config.MAX_WAIT)
                    print(f"   ↳ Sent! Waiting {wait}s …\n")
                    time.sleep(wait)
                else:
                    print("   ↳ Sent!\n")

            except Exception as exc:
                print(f"❌ Failed for SNo {sno} ({to_email}): {exc}")
                status_db[sno] = _record(sno, hr_name, to_email, company_name,
                                         NOT_SENT, error=str(exc))
                save_status_db(status_db)

        # ── 8. Verify newly sent emails via IMAP ───────────────────────────
        if sent_this_run:
            print(f"\n🔍 Verifying {len(sent_this_run)} newly sent emails via Gmail IMAP …")
            # Brief delay to allow Gmail to index the sent mail
            time.sleep(5)
            imap_results = verify_via_imap([x["email"] for x in sent_this_run])

            for item in sent_this_run:
                sno, email = item["sno"], item["email"]
                rec = status_db.get(sno, {})
                if imap_results.get(email):
                    status_db[sno] = {**rec, "status": SENT_VERIFIED,
                                      "verified_at": _now(), "updated_at": _now()}
                    print(f"   ✅ SENT_VERIFIED: SNo {sno} — {email}")
                else:
                    status_db[sno] = {**rec, "status": VERIFICATION_REQUIRED,
                                      "updated_at": _now()}
                    print(f"   ⚠️  VERIFICATION_REQUIRED: SNo {sno} — {email}")

            save_status_db(status_db)

    # ── 9. FINAL AUDIT REPORT ─────────────────────────────────────────────
    # Re-read Excel to always work from the current state (catches new rows)
    df_final   = load_excel()
    total_now  = len(df_final)
    status_db  = load_status_db()

    counts = {SENT_VERIFIED: 0, NOT_SENT: 0, VERIFICATION_REQUIRED: 0, SKIPPED: 0}
    attention = []   # rows that are NOT_SENT or VERIFICATION_REQUIRED

    for _, row in df_final.iterrows():
        sno = str(int(row["SNo"]))
        rec = status_db.get(sno, {})
        st  = rec.get("status", NOT_SENT)
        counts[st] = counts.get(st, 0) + 1

        if st in (NOT_SENT, VERIFICATION_REQUIRED):
            attention.append({
                "sno":    sno,
                "name":   rec.get("name",  row.get("Name", "Unknown")),
                "email":  rec.get("email", "Unknown"),
                "status": st,
            })

    reconciled = sum(counts.values())
    now_str    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    report = (
        f"\n{'='*64}\n"
        f"📋  FINAL AUDIT REPORT  —  {now_str}\n"
        f"{'='*64}\n"
        f"  Total recipients detected (Excel):   {total_now}\n"
        f"  ✅  Sent & Verified (SENT_VERIFIED):  {counts[SENT_VERIFIED]}\n"
        f"  ⚠️   Verification Required:            {counts[VERIFICATION_REQUIRED]}\n"
        f"  ❌  Not Sent (NOT_SENT):              {counts[NOT_SENT]}\n"
        f"  ⏭️   Skipped (SKIPPED):               {counts[SKIPPED]}\n"
        f"  {'─'*44}\n"
        f"  Reconciled Total:                    {reconciled} / {total_now} "
        f"{'✅' if reconciled == total_now else '❌ MISMATCH'}\n"
    )

    if attention:
        report += f"\n  🔎  Recipients requiring attention:\n"
        for a in attention:
            report += f"     SNo {a['sno']:>3}: {a['name'][:30]:<30} | {a['email']:<40} | {a['status']}\n"
    else:
        report += "\n  🎉  All detected recipients are either verified or intentionally skipped.\n"

    report += f"{'='*64}\n"

    print(report)

    # Save audit report to file
    with open(AUDIT_LOG, "a", encoding="utf-8") as f:
        f.write(report)

    # Write human-readable history.txt for manual verification
    write_history_txt(load_status_db())


if __name__ == "__main__":
    main()