import os
from dotenv import load_dotenv

# Load secrets from .env file when running locally.
load_dotenv()

# ==========================================
# CREDENTIALS & SECRETS
# ==========================================
YOUR_EMAIL = os.getenv("YOUR_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD")
YOUR_PHONE = os.getenv("YOUR_PHONE", "")

# ==========================================
# YOUR DETAILS
# ==========================================
YOUR_NAME = "Jelsinge Arun"
YOUR_LINKEDIN = "https://www.linkedin.com/in/arun-jelsinge-244949290/"
YOUR_GITHUB = "https://github.com/jelsingearun"
YOUR_PORTFOLIO = "https://jelsinge-arun-portfolio.vercel.app/"
YOUR_COURSE = "B.Tech in Information Technology (CGPA: 7.7/10)"
YOUR_COLLEGE = "Malla Reddy University, Hyderabad"

# ==========================================
# SYSTEM SETTINGS
# ==========================================
EXCEL_PATH = "hr_details.xlsx"
SHEET_NAME = "All Profiles"
RESUME_PATH = "Arun_Jelsinge_Resume.pdf"
SUBJECT = "Seeking Entry-Level Software Engineering Opportunity"

# Batch processing settings
BATCH_SIZE = 69       # How many emails to send every time you run app.py
MIN_WAIT = 15         # Minimum wait time between emails (seconds)
MAX_WAIT = 20         # Maximum wait time between emails (seconds)
SKIP_CONTACT_NAMES = {
    "Divya Vani",
    "Rishabh Mrinal",
    "Sudheer Jarajapu",
    "Sandhya Dodiya",
    "Rajashekar Manthena",
    "Saisudheer Jarajapu",
}
SKIP_CONTACT_EMAILS = {
    "divyarunja@gmail.com",
    "saisudheer662@gmail.com",
}

# ==========================================
# EMAIL TEMPLATE
# ==========================================
EMAIL_TEMPLATE = """Hi {HR_NAME},

Iam {YOUR_NAME}, a {YOUR_COURSE} graduate from {YOUR_COLLEGE}, and I’m looking for an opportunity to start my career as a Software Engineer or Full-Stack Developer.

I have hands-on experience building and testing software applications with Python, JavaScript, React, Node.js, FastAPI, MongoDB, and REST APIs. My work includes authentication, database integration, debugging, and automated and end-to-end testing.

A FEW THINGS I HAVE BUILT
01 · RoadCare — AI Road Damage Detection Platform
Built a full-stack road-damage reporting platform using React, Python/FastAPI, MongoDB, OpenCV, YOLOv8, and H3. Implemented JWT authentication, role-based access, input validation, rate limiting, and unit, integration, API, database, geospatial, and E2E testing.
02 · Academix — Academic Collaboration Platform
Developed a collaboration platform with React 19, JavaScript, Node.js, Express.js, MongoDB, TensorFlow.js, and JWT, including REST APIs, authentication, validation, persistence, and peer matching.
03 · GitHub Repo Automator
Created a GitHub automation CLI supporting folders, ZIP archives, and individual files, with secret scanning, sanitization, repository setup, logging, and dry-run execution.

WHY I’M REACHING OUT
I’m looking for an environment where I can contribute to real products, learn from experienced engineers, and take ownership of meaningful problems.

If {COMPANY_NAME} is hiring entry-level software engineers or interns—or if you know the right person on the engineering hiring team—I’d genuinely appreciate an introduction or referral.

One small ask: If my profile looks relevant, would you be open to pointing me toward the appropriate opportunity?

📄 Resume: Attached
💼 LinkedIn: {YOUR_LINKEDIN}
💻 GitHub: {YOUR_GITHUB}
🌐 Portfolio: {YOUR_PORTFOLIO}

Thank you for your time, {HR_NAME}. I know you probably receive a lot of messages like this, so I’ll keep this short—and let my work speak for itself.

Warm regards,
{YOUR_NAME}
{YOUR_COURSE} · {YOUR_COLLEGE}
📞 {YOUR_PHONE} · ✉️ {YOUR_EMAIL}
"""