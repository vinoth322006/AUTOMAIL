"""
AUTOMAIL Configuration Module
Loads settings from .env file and provides defaults.
"""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# ─── Email Settings ──────────────────────────────────────────────
GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ─── AI Settings (Templates are built-in, no API key needed) ────

# ─── Target Role ─────────────────────────────────────────────────
TARGET_ROLE = os.getenv("TARGET_ROLE", "AI/ML Engineer Intern")

# ─── File Paths ──────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HR_EXCEL_PATH = os.path.join(BASE_DIR, "HR Reachout List.xlsx")
DEFAULT_RESUME = os.getenv("DEFAULT_RESUME", "Vinoth_Resume.pdf")
DEFAULT_RESUME_PATH = os.path.join(BASE_DIR, DEFAULT_RESUME)
SEND_LOG_PATH = os.path.join(BASE_DIR, "send_log.csv")

# ─── Available Resumes ───────────────────────────────────────────
AVAILABLE_RESUMES = {
    "vinoth_main": {
        "filename": "VINOTH_P RESUME.pdf",
        "label": "Vinoth - Main Resume",
        "path": os.path.join(BASE_DIR, "VINOTH_P RESUME.pdf"),
    }
}

# ─── Rate Limiting ───────────────────────────────────────────────
DELAY_BETWEEN_EMAILS =4       # seconds between emails
BATCH_SIZE = 50                  # max emails per batch
MAX_DAILY_EMAILS = 450           # stay under Gmail's 500/day limit
MAX_PER_DOMAIN = 3               # max emails per company domain per batch (anti-spam guard)

# ─── Email Template Defaults ─────────────────────────────────────
DEFAULT_SUBJECT_TEMPLATE = "Application for {role} – {candidate_name}"
