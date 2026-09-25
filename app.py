import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import random
import os
import sys
import io

# Ensure UTF-8 encoding for Windows console output
if sys.stdout and sys.stdout.buffer:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if sys.stderr and sys.stderr.buffer:
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import config

HISTORY_FILE = "history.txt"

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

def send_email(to_email, subject, body):
    """Constructs and sends the email."""
    msg = MIMEMultipart()
    msg["From"] = config.YOUR_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    # Attach resume
    if os.path.exists(config.RESUME_PATH):
        with open(config.RESUME_PATH, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(config.RESUME_PATH))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(config.RESUME_PATH)}"'
            msg.attach(part)
    else:
        print(f"Warning: Resume not found at {config.RESUME_PATH}! Sending without attachment.")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(config.YOUR_EMAIL, config.APP_PASSWORD)
        server.send_message(msg)

def main():
    if not config.SENDING_ENABLED:
        print("Email sending is disabled in config.py. No messages were sent.")
        return

    if not config.YOUR_EMAIL or not config.APP_PASSWORD:
        print("Error: Email or App Password is missing in .env file.")
        sys.exit(1)

    try:
        df = pd.read_excel(config.EXCEL_PATH)
        if "Name" not in df.columns and "Personal Email" not in df.columns:
            if "Name" in df.iloc[0].values or "Personal Email" in df.iloc[0].values or "Email" in df.iloc[0].values:
                df = pd.read_excel(config.EXCEL_PATH, header=1)
        if "SNo" not in df.columns:
            df["SNo"] = range(1, len(df) + 1)
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
        hr_name = row.get("Name") if pd.notna(row.get("Name")) and str(row.get("Name")).strip() else "Hiring Manager"
        company_val = row.get("Company Name") if pd.notna(row.get("Company Name")) and str(row.get("Company Name")).strip() else row.get("Company")
        company_name = company_val if pd.notna(company_val) and str(company_val).strip() else "your company"

        # Prioritize Personal Email column
        to_email = row.get("Personal Email")
        if pd.isna(to_email) or not str(to_email).strip():
            to_email = row.get("Email")

        if pd.notna(to_email):
            to_email = str(to_email).split(";")[0].strip()

        if not to_email or pd.isna(to_email):
            print(f"⚠️ Skipping SNo {sno} - No email address provided.")
            update_history(sno)
            continue

        body = create_email_body(hr_name, company_name)

        print(f"[{sno}] Sending email to {hr_name} ({to_email}) at {company_name}...")
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