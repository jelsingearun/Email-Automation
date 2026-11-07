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
YOUR_COURSE = "B.Tech in Information Technology"
YOUR_COLLEGE = "Malla Reddy University, Hyderabad"

# ==========================================
# SYSTEM SETTINGS
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "hr_details.xlsx")
SHEET_NAME = "Sheet1"
RESUME_PATH = os.path.join(BASE_DIR, "Jelsinge Arun Resume.pdf")
SUBJECT = f"Quick intro — {YOUR_NAME}"

# Batch processing settings
# Each run sends at most BATCH_SIZE contacts.
BATCH_SIZE = int(os.getenv("BATCH_SIZE", 20))
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
"saihasini1@gmail.com",
"shekarmanthena603@gmail.com",
"jaishrees@google.com",
"sandhyazaverdodiya@gmail.com",
"rajat.pandit016@gmail.com",
"veerabhadrahareesh@gmail.com",
"ashrithugadi@gmail.com",
"sayantann2002@gmail.com",
"nagulaanusha21@gmail.com",
"iamsumit.sc@gmail.com",
"tarasha2912@gmail.com",
"jinnajhanvi@gmail.com",
"saigoudsai041@gmail.com",
"deepasm36@gmail.com",
"tulikapatra1998@gmail.com",
"naimisha.rao@gmail.com",
"corporatetrg@aptech.co.in",
"rajeswari.pamoti@gmail.com",
"stummala@planittesting.com",
"divya.vullamparthi@pragmaedge.com",
"mounikapoosarla3@gmail.com",
"info@peakxv.com",
"info@locofast.com",
"info@inito.com",
"careers@pocketfm.com",
"careers@aarnabio.com",
"hello@ecofueltech.in",
"hello@custona.com",
"hello@bluesemi.io"

}

# ==========================================
# EMAIL TEMPLATES (Persona-Tailored)
# ==========================================

# 1. Founders, CEOs, CTOs & Executive Leadership
TEMPLATE_FOUNDER = """Hi {HR_NAME},

I came across what you're building at {COMPANY_NAME} and wanted to reach out directly.

A bit about me: I'm a software engineer graduating this year with a {YOUR_COURSE} from {YOUR_COLLEGE}. Rather than just wrapping APIs or writing classroom code, I focus on building production-facing systems where performance, clean architecture, and testing aren't afterthoughts.

A couple of projects I've built and shipped:
• RoadCare (https://github.com/jelsingearun/RoadCare---AI-Based-Road-Damage-Pothole-Detection-System) — Full-stack detection platform using Python/FastAPI, React, MongoDB, and YOLOv8 with JWT auth, rate limiting, and 6 layers of automated testing.
• GitHub Repo Automator (https://github.com/jelsingearun/github-repo-automator) — Developer CLI supporting multi-format inputs with automated secret scanning (detecting private keys and high-entropy credentials) and dry-run execution.

My stack maps naturally to product engineering in Python, React, and REST APIs. I understand early-stage engineering needs builders who don't need hand-holding and can take real ownership from day one.

More about me: {YOUR_PORTFOLIO}

If {COMPANY_NAME} has room for an early-career engineer or intern who ships fast and writes reliable code, I think we'd both get something out of a quick conversation.

Best regards,
Arun
{YOUR_EMAIL} · {YOUR_PHONE}
GitHub: {YOUR_GITHUB} | LinkedIn: {YOUR_LINKEDIN}
"""

# 2. Software Engineers, Tech Leads & Technical Peers (Referral ask)
TEMPLATE_ENGINEER = """Hi {HR_NAME},

I came across your work and what the engineering team is building at {COMPANY_NAME}.

A bit about me: I'm an entry-level software engineer graduating this year with a {YOUR_COURSE} from {YOUR_COLLEGE}. My core stack is Python, FastAPI, React, and MongoDB, with a big emphasis on clean API design, authentication, and test coverage (unit, integration, and E2E).

A couple of things I've built recently:
• RoadCare (https://github.com/jelsingearun/RoadCare---AI-Based-Road-Damage-Pothole-Detection-System) — Full-stack road damage detection platform with Python/FastAPI, React, and YOLOv8, tested across database, geospatial, and API layers.
• GitHub Repo Automator (https://github.com/jelsingearun/github-repo-automator) — CLI tool for repo initialization, template generation, and automated secret scanning before code reaches remote repos.

More about me: {YOUR_PORTFOLIO}

I really respect what your team is building at {COMPANY_NAME}. If you think my technical foundation looks solid and your team is hiring, would you be open to pointing me toward the right engineering manager or submitting an internal referral?

Either way, appreciate you taking a look at my work.

Best regards,
Arun
{YOUR_EMAIL} · {YOUR_PHONE}
GitHub: {YOUR_GITHUB} | LinkedIn: {YOUR_LINKEDIN}
"""

# 3. HR, Talent Acquisition & Recruiters (Role matching & credentials)
TEMPLATE_HR = """Hi {HR_NAME},

I'm reaching out because I've been following {COMPANY_NAME} and wanted to check if your team is hiring for entry-level Software Engineers or Full-Stack Developer roles.

A bit about my background: I'm graduating this year with a {YOUR_COURSE} from {YOUR_COLLEGE}. My technical foundation centers on Python, FastAPI, React, Node.js, and MongoDB, with hands-on experience building REST APIs, role-based auth, and automated testing suites.

A couple of recent projects:
• RoadCare (https://github.com/jelsingearun/RoadCare---AI-Based-Road-Damage-Pothole-Detection-System) — Full-stack AI road inspection system using Python/FastAPI, React, MongoDB, and YOLOv8 with end-to-end testing.
• GitHub Repo Automator (https://github.com/jelsingearun/github-repo-automator) — Automation CLI featuring credential/secret scanning and automated repository setup.

My resume is attached, and you can see more of my code and live work here:
Portfolio: {YOUR_PORTFOLIO}
GitHub: {YOUR_GITHUB}

If you have an open junior developer role or an upcoming hiring drive at {COMPANY_NAME}, I'd welcome the chance to speak with your hiring team.

Best regards,
Arun
{YOUR_EMAIL} · {YOUR_PHONE}
LinkedIn: {YOUR_LINKEDIN}
"""

# 4. General Enquiries, Product & Other Roles
TEMPLATE_GENERAL = """Hi {HR_NAME},

I'm reaching out to introduce myself to the team at {COMPANY_NAME}.

I'm a graduating software engineer with a {YOUR_COURSE} from {YOUR_COLLEGE}, specializing in Python, FastAPI, React, and MongoDB. I build production-ready applications with an emphasis on clean architecture, security, and test-driven development.

A couple of projects I've recently shipped:
• RoadCare (https://github.com/jelsingearun/RoadCare---AI-Based-Road-Damage-Pothole-Detection-System) — Full-stack road damage detection app (Python/FastAPI, React, YOLOv8).
• GitHub Repo Automator (https://github.com/jelsingearun/github-repo-automator) — CLI utility with built-in secret scanning and automated repo scaffolding.

Portfolio: {YOUR_PORTFOLIO} | GitHub: {YOUR_GITHUB}

If {COMPANY_NAME} is looking for entry-level software engineering talent, could you please point me toward the appropriate hiring manager or engineering lead?

Thank you for your time and assistance.

Best regards,
Arun
{YOUR_EMAIL} · {YOUR_PHONE}
LinkedIn: {YOUR_LINKEDIN}
"""

EMAIL_TEMPLATES = {
    "FOUNDER": TEMPLATE_FOUNDER,
    "ENGINEER": TEMPLATE_ENGINEER,
    "HR": TEMPLATE_HR,
    "GENERAL": TEMPLATE_GENERAL,
}

# Default fallback template for backwards compatibility
EMAIL_TEMPLATE = TEMPLATE_HR