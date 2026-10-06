# Cold Email Automator

A Python script to automate sending personalized cold emails with resume attachments in calculated, randomized batches to avoid spam detection. 

## Features
- **Batch Processing:** Sends emails in batches (e.g., 50 at a time).
- **Auto-Save & Resume:** Saves the last successfully sent serial number (`SNo`) in a `history.txt` file. If the program crashes or you stop it manually, the next run will pick up exactly where it left off.
- **Anti-Spam Delay:** Pauses for a randomized time (15–40 seconds) between each email.
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
* `BATCH_SIZE`: How many emails to send per script execution.
* Review `EMAIL_TEMPLATE` and make sure it aligns with your preferences.

### 4. Prepare the Excel File

The project expects a local Excel workbook configured by `EXCEL_PATH` and the
`All Profiles` sheet by default. The workbook is intentionally not included in
this repository; provide your own contact file locally and do not commit it.
The column headings are on the second row of the sheet.
**Required Column Names:**

* `Name` (HR's Name)
* `Company Name`
* `Corporate Email` or `Personal Email`

Serial numbers are generated automatically from the spreadsheet row order.
Corporate email is preferred, with personal email used as a fallback.
The configured skip list excludes Divya Vani (`divyarunja@gmail.com`),
Rishabh Mrinal, and Sudheer Jarajapu from sending.

### 5. Add Your Resume

Place `Arun_Jelsinge_Resume.pdf` in the same folder (or modify `RESUME_PATH` in `config.py`).

## How to Use

Simply run the script:

```bash
python app.py
```

* **First Run:** It will send emails for `SNo` 1 up to your `BATCH_SIZE` (e.g., 1 to 50). It will create `history.txt` and save `50`.
* **Second Run:** Tomorrow (or whenever you run it again), it reads `history.txt` and automatically sends `SNo` 51 to 100.
* **Resetting:** If you want to start from the beginning again, simply delete `history.txt` or change the number inside it to `0`.
