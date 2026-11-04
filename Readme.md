# Cold Email Automator

A local Python utility for sending personalized job-search emails from an Excel
workbook. It uses Gmail SMTP with an App Password, sends a bounded batch per
run, and records progress only after a row has been sent or intentionally
skipped.

## Features
- **Batch Processing:** Sends at most `BATCH_SIZE` contacts per run.
- **Auto-Save & Resume:** Atomically saves the last processed serial number (`SNo`) in `history.txt`.
- **Safe Preview:** `--dry-run` validates the workbook and previews recipients without sending or changing history.
- **Reusable SMTP session:** Authenticates once per batch instead of once per recipient.
- **Configurable delay:** Pauses for a randomized time between emails.
- **Dynamic Variables:** Customizes the `{HR_NAME}` and `{COMPANY_NAME}` inside the email body for each row in your Excel sheet.

## Setup Instructions

### 1. Install Dependencies
You need Python installed on your system. Run the following command to install the required libraries:
```bash
python -m venv venv
venv\Scripts\activate
python -m pip install -r requirements.txt
```
For Mac/Linux:
```bash
python3 -m venv venv
source venv/bin/activate
python3 -m pip install -r requirements.txt
```

### 2. Configure Environment Variables (`.env`)

Copy `.env.example` to `.env` in the main folder and add your Gmail credentials.
**Note:** Do not use your standard Gmail password. You must generate an **App Password**.

```env
YOUR_EMAIL=your-email@gmail.com
APP_PASSWORD=xxxx xxxx xxxx xxxx
YOUR_PHONE=your-phone-number
```

*(To get an App Password: Go to Google Account > Security > 2-Step Verification > App Passwords > Generate a password for "Mail").*

The `.env` file is local-only and is excluded from version control. Never commit
real credentials.

### 3. Setup Configurations & Template (`config.py`)

Open `config.py` and modify your personal details:

* `YOUR_NAME`, `YOUR_PHONE`, `YOUR_LINKEDIN`, etc.
* `BATCH_SIZE`: Maximum number of contacts to process per execution.
* Review `EMAIL_TEMPLATE` and make sure it aligns with your preferences.

### 4. Prepare the Excel File

The project expects a local Excel workbook configured by `EXCEL_PATH` and the
`Sheet1` sheet by default. The workbook is intentionally not included in this
repository; provide your own contact file locally and do not commit it. The
loader accepts column headings on either the first or second row of the sheet.
**Required Column Names:**

* `Name` (HR's Name)
* `Company Name`
* `Corporate Email` or `Personal Email`

Serial numbers are generated automatically from the spreadsheet row order.
Corporate email is preferred, with personal email used as a fallback.
The configured skip list excludes Divya Vani (`divyarunja@gmail.com`),
Rishabh Mrinal, and Sudheer Jarajapu from sending.

### 5. Add Your Resume

Place `Jelsinge Arun Resume.pdf` in the same folder (or modify `RESUME_PATH`
in `config.py`). The attachment is optional; the script sends without one if it
is missing.

## How to Use

Preview the next batch without sending:

```bash
python app.py --dry-run
```

Send the next configured batch:

```bash
python app.py
```

Use `--limit N` to override the configured batch size for one run:

```bash
python app.py --limit 5
```

* **First Run:** It processes `SNo` 1 through `BATCH_SIZE`.
* **Next Run:** It resumes after the last processed `SNo`.
* **Failed Send:** The failed row is not marked as processed, so it is retried
  on the next run.
* **Resetting:** Delete `history.txt` or replace its value with `0` only when
  you intentionally want to restart.

The script does not guarantee delivery or bypass provider limits. Use it only
for contacts who have a legitimate reason to receive your message and respect
applicable email and privacy laws.
