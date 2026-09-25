import os
from dotenv import load_dotenv

# Load secrets from .env file
load_dotenv()

# ==========================================
# CREDENTIALS & SECRETS
# ==========================================
YOUR_EMAIL = os.getenv("YOUR_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD", "").replace(" ", "") if os.getenv("APP_PASSWORD") else None
YOUR_PHONE = os.getenv("YOUR_PHONE")

# ==========================================
# ==========================================
# YOUR DETAILS
# ==========================================
YOUR_NAME = "Arun Jelsinge"
YOUR_LINKEDIN = "linkedin.com/in/arun-jelsinge-244949290/"
YOUR_GITHUB = "github.com/jelsingearun"
YOUR_PORTFOLIO = "jelsinge-arun-portfolio.vercel.app/"
YOUR_COURSE = "B.Tech in Information Technology (7.7 CGPA)"
YOUR_COLLEGE = "Malla Reddy University, Hyderabad"

# ==========================================
# SYSTEM SETTINGS
# ==========================================
EXCEL_PATH = "hr_details.xlsx"
RESUME_PATH = "Arun_Jelsinge_Resume.pdf"
SUBJECT = "Seeking Entry-Level Software Engineering Opportunity"

# Batch processing settings
SENDING_ENABLED = True  # Set to True only when you explicitly want to send mail.
BATCH_SIZE = 50       # How many emails to send every time you run app.py
MIN_WAIT = 15         # Minimum wait time between emails (seconds)
MAX_WAIT = 40         # Maximum wait time between emails (seconds)

# ==========================================
# EMAIL TEMPLATE
# ==========================================
EMAIL_TEMPLATE = """Hi {HR_NAME},

I’m {YOUR_NAME}, a {YOUR_COURSE} graduate from {YOUR_COLLEGE}, and I’m currently looking for an opportunity to start my career as a Software Engineer / Full-Stack Developer.

Rather than sending a generic application, I wanted to share a quick snapshot of what I can actually build:

WHAT I BUILD
→ Full-Stack Applications — React 19, Node.js, Express, Python, FastAPI, MongoDB
→ AI/ML & Computer Vision — Python, TensorFlow.js, OpenCV, YOLOv8
→ Automation & Security Tools — Git, GitHub CLI, Security & Secret Scanning
→ Software Testing — Unit, Integration, API, & E2E Testing (pytest, Jest)

A FEW THINGS I'VE BUILT
01 · RoadCare — AI Road Damage Detection Platform
Full-stack platform with React, Python/FastAPI, MongoDB, OpenCV, and YOLOv8 featuring JWT authentication and role-based access.
02 · Academix — Academic Collaboration Platform
Student collaboration app with React 19, Node.js, Express, MongoDB, and TensorFlow.js for compatibility-based peer matching.
03 · GitHub Repo Automator
CLI tool in Python & GitHub CLI for automated repository creation with security scanning across threat categories.

WHY I’M REACHING OUT
I’m looking for an environment where I can contribute to real products, learn from experienced engineers, and take ownership of meaningful problems.

If {COMPANY_NAME} is hiring entry-level engineers/interns—or if you know the right person on the engineering hiring team—I’d genuinely appreciate an introduction or referral.

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