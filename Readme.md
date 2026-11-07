# Cold Email Automator

A Python automation engine designed for high-conversion, personalized job outreach from an Excel workbook. It connects via Gmail SMTP with an App Password, features **automatic persona-based message routing** (Founders, Engineers, HR, General), enforces safe batch limits with anti-spam delays, and tracks progress atomically.

---

## Key Features

- **Persona-Tailored Outreach:** Automatically inspects the recipient's `Job Title` and `Profile Type` in Excel to route to one of four specialized email templates:
  - **Founders & Leadership (`FOUNDER`):** Focuses on builder velocity, direct ownership, and solving problems without handholding.
  - **Software Engineers & Tech Leads (`ENGINEER`):** Peer-to-peer technical respect highlighting stack alignment and asking for internal team referrals.
  - **HR & Talent Acquisition (`HR`):** Clear junior/entry-level role matching, immediate availability, and resume attachment.
  - **General Inquiries (`GENERAL`):** Polite, concise inquiry asking to route the message to technical hiring.
- **Cold Email Writer Standard:** Replaces corporate cover letter fluff with human, collegial outreach (~190 words), natural first-name greetings (`Hi [Name],` / `Hi there,`), live project links, and confident, low-pressure calls to action.
- **Batch Processing & Rate Limiting:** Sends at most `BATCH_SIZE` contacts per run with randomized delays (`MIN_WAIT` to `MAX_WAIT` seconds) between emails to respect Gmail sending limits.
- **Safe Dry Run Preview:** `--dry-run` previews the next batch of recipients, company names, and detected personas without connecting to SMTP or changing history.
- **Auto-Save & Resume:** Atomically persists the last processed serial number (`SNo`) in `history.txt`. If an email fails mid-batch, progress is saved up to the last successful send.
- **Reusable SMTP Session:** Authenticates with Gmail once per batch session rather than reconnecting per recipient.
- **In-Memory Caching:** Loads the resume PDF and spreadsheet headers once per execution.

---

## Setup Instructions

### 1. Install Dependencies

```bash
python -m venv venv
venv\Scripts\activate      # Windows
# source venv/bin/activate # Mac/Linux

pip install -r requirements.txt
```

### 2. Configure Environment Variables (`.env`)

Copy `.env.example` to `.env` in the root folder and add your Gmail credentials:

```env
YOUR_EMAIL=your-email@gmail.com
APP_PASSWORD=xxxx xxxx xxxx xxxx
YOUR_PHONE=your-phone-number
```

> **Note:** Generate an **App Password** via Google Account > Security > 2-Step Verification > App Passwords (choose "Mail"). Never commit real credentials.

### 3. Customize Details & Settings (`config.py`)

Open `config.py` to customize:
- `YOUR_NAME`, `YOUR_PHONE`, `YOUR_PORTFOLIO`, `YOUR_GITHUB`, `YOUR_LINKEDIN`
- `BATCH_SIZE`: Maximum contacts sent per run (default: `20`)
- `MIN_WAIT` / `MAX_WAIT`: Wait interval in seconds between successive emails (default: `15`–`20`s)
- `SKIP_CONTACT_NAMES` / `SKIP_CONTACT_EMAILS`: Blacklist exclusions

### 4. Prepare the Excel File (`hr_details.xlsx`)

The script expects an Excel workbook configured via `EXCEL_PATH` (`hr_details.xlsx`) and `SHEET_NAME` (`Sheet1`).

Supported columns (normalized automatically):
- `Name`: Recruiter, Engineer, or Founder name
- `Job Title` / `Designation`: Used for automatic persona detection
- `Company Name` / `Company`: Target company name
- `Corporate Email` or `Personal Email`: Primary recipient email
- `Profile Type`: (Optional) `HR`, `Non-HR`, or `Founder`

### 5. Add Your Resume

Place `Arun_Jelsinge_Resume.pdf` in the project root (or update `RESUME_PATH` in `config.py`). The script automatically attaches the PDF if present.

---

## How to Use

### Preview Next Batch (Dry Run)
Inspect the next contacts and detected personas without sending emails:
```bash
python app.py --dry-run
```

### Send Next Batch
Dispatches the next configured batch (default 20 emails):
```bash
python app.py
```

### Override Batch Limit
Send a smaller or larger batch for a single run:
```bash
python app.py --limit 5
```

---

## Progress Tracking & Resetting

- **Progress:** Stored in `history.txt` as the last completed `SNo`. Each subsequent run starts at `SNo = last_sent + 1`.
- **Restarting:** To restart from the beginning, set the value in `history.txt` to `0`.
- **Targeted Run:** To resume from a specific row `N`, set `history.txt` to `N - 1`.

---
