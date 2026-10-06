import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import random
import os
import sys
import re

import config

HISTORY_FILE = "history.txt"

def normalize_contact_columns(df):
    """Map common spreadsheet header variants to the names used by the app."""
    normalized = {}
    aliases = {
        "SNo": {"sno", "sno.", "s.no", "s no", "serialno", "serial no", "serialnumber", "serial number"},
        "Name": {"name", "hrname", "hr name", "hiring manager", "hiring manager name"},
        "Company": {"company", "company name", "organization", "organisation"},
        "Email": {"email", "email address", "e-mail", "e-mail address"},
        "Personal Email": {"personal email", "personal email address"},
        "Corporate Email": {"corporate email", "corporate email address", "work email", "official email"},
    }

    for column in df.columns:
        key = " ".join(str(column).strip().lower().replace("_", " ").split())
        key_without_punctuation = (
            key.replace(".", "").replace("-", "").replace(" ", "")
        )
        for target, candidates in aliases.items():
            candidate_keys = {
                item.replace(".", "").replace("-", "").replace(" ", "")
                for item in candidates
            }
            if key in candidates or key_without_punctuation in candidate_keys:
                normalized[column] = target
                break

    df = df.rename(columns=normalized)
    if "Email" not in df.columns:
        email_columns = [
            column
            for column in ("Corporate Email", "Personal Email")
            if column in df.columns
        ]
        if email_columns:
            df["Email"] = df[email_columns[0]]
            for column in email_columns[1:]:
                df["Email"] = df["Email"].fillna(df[column])
                df["Email"] = df["Email"].replace("", pd.NA).fillna(df[column])

    if "SNo" not in df.columns:
        df.insert(0, "SNo", range(1, len(df) + 1))

    required_columns = {"Email"}
    missing_columns = required_columns - set(df.columns)
    if missing_columns:
        available_columns = ", ".join(str(column) for column in df.columns)
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(
            f"Missing required column(s): {missing}. "
            f"Available columns: {available_columns}"
        )

    return df

def load_contacts():
    """Load the contact sheet and normalize its grouped-header layout."""
    df = pd.read_excel(
        config.EXCEL_PATH,
        sheet_name=config.SHEET_NAME,
        header=1,
    )
    return normalize_contact_columns(df)

def get_last_sent_sno():
    """Reads the history file to find the last successfully sent SNo."""
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            content = f.read().strip()
            if content.isdigit():
                return int(content)
    return 0  # Default to 0 if no history exists

def update_history(sno):
    """Saves the latest successfully sent SNo to the history file."""
    with open(HISTORY_FILE, "w") as f:
        f.write(str(sno))

def create_email_body(hr_name, company_name):
    """Formats the email template with actual variables."""
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
        YOUR_EMAIL=config.YOUR_EMAIL
    )

def _normalized_emails(value):
    """Return normalized addresses from a cell that may contain multiple emails."""
    if pd.isna(value):
        return set()
    return {
        item.strip().casefold()
        for item in str(value).replace(";", ",").split(",")
        if item.strip()
    }

def get_recipient_emails(value):
    """Return valid individual email addresses from a spreadsheet cell."""
    candidates = str(value).replace(";", ",").split(",")
    valid_emails = []
    for candidate in candidates:
        email = candidate.strip()
        if email and re.fullmatch(r"[^@\s<>;,]+@[^@\s<>;,]+\.[^@\s<>;,]+", email):
            valid_emails.append(email)
    return valid_emails

def should_skip_contact(hr_name, *email_values):
    """Return whether a contact is configured to be excluded from sending."""
    name = str(hr_name).strip().casefold()
    skip_names = {value.strip().casefold() for value in config.SKIP_CONTACT_NAMES}
    skip_emails = {value.strip().casefold() for value in config.SKIP_CONTACT_EMAILS}
    contact_emails = set()
    for value in email_values:
        contact_emails.update(_normalized_emails(value))
    return name in skip_names or bool(contact_emails & skip_emails)

def send_email(to_email, subject, body):
    """Constructs and sends the email."""
    recipient_emails = get_recipient_emails(to_email)
    if not recipient_emails:
        raise ValueError(f"No valid recipient email found in: {to_email!r}")

    for recipient in recipient_emails:
        if should_skip_contact("", recipient):
            print(f"⏭️ Skipping blocked recipient {recipient}.")
            continue

        msg = MIMEMultipart()
        msg["From"] = config.YOUR_EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        if os.path.exists(config.RESUME_PATH):
            with open(config.RESUME_PATH, "rb") as f:
                part = MIMEApplication(
                    f.read(),
                    Name=os.path.basename(config.RESUME_PATH),
                )
                part["Content-Disposition"] = (
                    f'attachment; filename="{os.path.basename(config.RESUME_PATH)}"'
                )
                msg.attach(part)
        else:
            print(
                f"Warning: Resume not found at {config.RESUME_PATH}! "
                "Sending without attachment."
            )

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(config.YOUR_EMAIL, config.APP_PASSWORD)
            server.send_message(msg)

def main():
    if not config.YOUR_EMAIL or not config.APP_PASSWORD:
        print("Error: Email or App Password is missing in .env file.")
        sys.exit(1)

    try:
        df = load_contacts()
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        sys.exit(1)

    last_sent = get_last_sent_sno()
    start_sno = last_sent + 1
    end_sno = last_sent + config.BATCH_SIZE

    # Filter Excel rows to match the batch we want to send
    batch_df = df[(df["SNo"] >= start_sno) & (df["SNo"] <= end_sno)]

    if batch_df.empty:
        print(f"\n✅ All caught up! No more contacts to email. (Last sent SNo was {last_sent})")
        print("If you added new contacts to the Excel sheet, make sure they have a valid SNo.")
        return

    print(f"\n🚀 Starting batch: Sending emails for SNo {start_sno} to {start_sno + len(batch_df) - 1}...")

    for index, row in batch_df.iterrows():
        sno = row.get("SNo")
        hr_name = row.get("Name", "Hiring Manager")
        to_email = row.get("Email")
        company_name = row.get("Company", "your company")

        if should_skip_contact(
            hr_name,
            to_email,
            row.get("Personal Email"),
            row.get("Corporate Email"),
        ):
            print(f"⏭️ Skipping {hr_name} ({to_email}) - configured skip contact.")
            update_history(sno)
            continue

        if pd.isna(to_email):
            print(f"⚠️ Skipping SNo {sno} - No email address provided.")
            update_history(sno)
            continue

        body = create_email_body(hr_name, company_name)

        recipients = get_recipient_emails(to_email)
        print(
            f"[{sno}] Sending email to {hr_name} "
            f"({', '.join(recipients) or to_email}) at {company_name}..."
        )
        try:
            send_email(to_email, config.SUBJECT, body)
            
            # Save progress immediately after successful send
            update_history(sno) 
            
            # Random wait between emails (except for the very last one in the batch)
            if sno != batch_df["SNo"].max():
                wait_time = random.randint(config.MIN_WAIT, config.MAX_WAIT)
                print(f"   ↳ Success! Waiting {wait_time} sec to avoid spam filters...\n")
                time.sleep(wait_time)
            else:
                print("   ↳ Success!\n")

        except Exception as e:
            print(f"❌ Failed to send to {to_email}. Error: {e}")
            print(f"🛑 Terminating batch early. Last successful SNo: {get_last_sent_sno()}")
            break

    print(f"\n🎉 Batch process completed! Total emails sent so far: {get_last_sent_sno()}")

if __name__ == "__main__":
    main()