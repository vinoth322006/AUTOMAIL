"""
AUTOMAIL - Automated Cold Email System with AI-Generated Templates
==================================================================
Reads HR contacts from Excel, uses pre-generated AI templates for
personalized cold emails, and sends them via Gmail SMTP.

Templates were generated using Google Gemini AI via Proxima.
No API key required - templates are built-in with smart rotation.

Usage:
    python automail.py --help
    python automail.py --dry-run --count 3          # Preview 3 emails
    python automail.py --test                        # Send test to yourself
    python automail.py --start 1 --count 50          # Send batch of 50
    python automail.py --list-resumes                # Show available resumes
    python automail.py --resume vinoth_skct --count 10
"""

import argparse
import csv
import io
import logging
import os
import re
import smtplib
import sys
import time
import traceback
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import make_msgid, formatdate
from hashlib import md5
import imaplib
import email

import openpyxl
from PyPDF2 import PdfReader

import config

# Fix Windows console encoding for Unicode
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ═══════════════════════════════════════════════════════════════════
# DEBUG LOGGING SYSTEM
# ═══════════════════════════════════════════════════════════════════

DEBUG_MODE = False
EMAIL_REGEX = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')

# File logger — always writes to debug.log
_file_logger = logging.getLogger("automail")
_file_logger.setLevel(logging.DEBUG)
_log_path = os.path.join(config.BASE_DIR, "debug.log")
_fh = logging.FileHandler(_log_path, encoding="utf-8")
_fh.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)-7s] %(funcName)-20s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
))
_file_logger.addHandler(_fh)
_file_logger.info("=" * 60)
_file_logger.info("AUTOMAIL session started")


def debug(msg):
    """Print to console only if --debug flag is set. Always logs to file."""
    _file_logger.debug(msg)
    if DEBUG_MODE:
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"  \033[90m[{ts}] DBG  {msg}\033[0m")


def validate_email(email):
    """Validate email format. Returns (is_valid, reason)."""
    if not email:
        return False, "empty"
    if not isinstance(email, str):
        return False, f"not a string: {type(email)}"
    if not EMAIL_REGEX.match(email):
        return False, f"bad format: {email}"
    if len(email) > 254:
        return False, f"too long ({len(email)} chars)"
        
    # Block system/bounce emails like "Mail Delivery Subsystem"
    email_lower = email.lower()
    system_keywords = ["mailer-daemon", "postmaster", "noreply", "no-reply", "delivery"]
    for keyword in system_keywords:
        if keyword in email_lower:
            return False, f"system/bounce email blocked: {email}"
            
    return True, "ok"


def health_check():
    """Run system diagnostics and report issues."""
    print(f"\n  {Colors.BOLD}--- AUTOMAIL Health Check ---{Colors.RESET}\n")
    issues = []

    # .env file
    env_path = os.path.join(config.BASE_DIR, ".env")
    _chk(os.path.exists(env_path), ".env file exists", issues, "Copy .env.example to .env")

    # Gmail
    _chk(bool(config.GMAIL_ADDRESS), f"Gmail: {config.GMAIL_ADDRESS[:4]}***" if config.GMAIL_ADDRESS else "Gmail address set", issues, "Set GMAIL_ADDRESS in .env")
    _chk(bool(config.GMAIL_APP_PASSWORD), "Gmail App Password set", issues, "Set GMAIL_APP_PASSWORD in .env — get it at myaccount.google.com/apppasswords")

    # Excel
    _chk(os.path.exists(config.HR_EXCEL_PATH), f"Excel file: {os.path.basename(config.HR_EXCEL_PATH)}", issues, f"Missing: {config.HR_EXCEL_PATH}")

    # Resumes
    for key, info in config.AVAILABLE_RESUMES.items():
        _chk(os.path.exists(info["path"]), f"Resume [{key}]: {info['filename']}", issues)

    # Templates
    _chk(len(EMAIL_TEMPLATES) > 0, f"Email templates: {len(EMAIL_TEMPLATES)} loaded", issues, "EMAIL_TEMPLATES is empty!")
    for i, t in enumerate(EMAIL_TEMPLATES):
        _chk("{company}" in t["body"] and "{hr_name}" in t["body"], f"  Template {i+1} placeholders valid", issues, f"Template {i+1} missing placeholders")

    # Log file
    _chk(True, f"Debug log: {_log_path}", issues)

    # Send log
    if os.path.exists(config.SEND_LOG_PATH):
        sent = get_sent_emails()
        print(f"  {Colors.GREEN}+{Colors.RESET}  Send log: {len(sent)} emails sent previously")
    else:
        print(f"  {Colors.DIM}~{Colors.RESET}  No send log yet (first run)")

    # SMTP test
    if config.GMAIL_ADDRESS and config.GMAIL_APP_PASSWORD:
        try:
            with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT, timeout=10) as s:
                s.starttls()
                s.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
            print(f"  {Colors.GREEN}+{Colors.RESET}  SMTP login: SUCCESS")
        except smtplib.SMTPAuthenticationError:
            issues.append("SMTP login FAILED — wrong App Password")
            print(f"  {Colors.RED}x{Colors.RESET}  SMTP login: AUTH FAILED")
        except Exception as e:
            issues.append(f"SMTP error: {e}")
            print(f"  {Colors.RED}x{Colors.RESET}  SMTP connection: {e}")

    # Summary
    print()
    if issues:
        print(f"  {Colors.RED}{Colors.BOLD}{len(issues)} issue(s) found:{Colors.RESET}")
        for iss in issues:
            print(f"    {Colors.RED}>{Colors.RESET} {iss}")
    else:
        print(f"  {Colors.GREEN}{Colors.BOLD}All checks passed!{Colors.RESET}")
    print()
    return len(issues) == 0


def _chk(condition, label, issues, fix_hint=None):
    if condition:
        print(f"  {Colors.GREEN}+{Colors.RESET}  {label}")
    else:
        msg = f"{label} — FAILED" + (f" ({fix_hint})" if fix_hint else "")
        print(f"  {Colors.RED}x{Colors.RESET}  {msg}")
        issues.append(msg)


# ═══════════════════════════════════════════════════════════════════
# COLORS & FORMATTING
# ═══════════════════════════════════════════════════════════════════

class Colors:
    HEADER = "\033[95m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RESET = "\033[0m"


def banner():
    print(f"""{Colors.CYAN}{Colors.BOLD}
    =====================================================
       AUTOMAIL - AI Cold Email Automation System
    ====================================================={Colors.RESET}
    """)


def log(msg, level="info"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    icons = {
        "info":    f"{Colors.BLUE}i{Colors.RESET}",
        "success": f"{Colors.GREEN}+{Colors.RESET}",
        "warning": f"{Colors.YELLOW}!{Colors.RESET}",
        "error":   f"{Colors.RED}x{Colors.RESET}",
        "send":    f"{Colors.GREEN}>{Colors.RESET}",
        "ai":      f"{Colors.CYAN}*{Colors.RESET}",
        "wait":    f"{Colors.DIM}~{Colors.RESET}",
    }
    icon = icons.get(level, icons["info"])
    print(f"  {Colors.DIM}[{timestamp}]{Colors.RESET} {icon}  {msg}")


# ═══════════════════════════════════════════════════════════════════
# EMAIL TEMPLATES SYSTEM
# ═══════════════════════════════════════════════════════════════════

_TEMPLATES_PATH = os.path.join(config.BASE_DIR, "templates.json")

def load_templates():
    """Load templates from templates.json or return defaults."""
    if os.path.exists(_TEMPLATES_PATH):
        try:
            with open(_TEMPLATES_PATH, "r", encoding="utf-8") as f:
                import json
                return json.load(f)
        except Exception as e:
            print(f"Error loading templates: {e}")
    
    # Defaults if file missing or corrupt
    return [
        {
            "subject": "Application for {role} – Vinoth Palanivel",
            "body": "Dear {hr_name},\n\nI am writing to submit my application for the {role} position at {company}..."
        }
    ]

def save_templates(templates):
    """Save templates to templates.json."""
    import json
    with open(_TEMPLATES_PATH, "w", encoding="utf-8") as f:
        json.dump(templates, f, indent=2)

def render_template(template, contact, role):
    """Render a template with contact details."""
    body = template["body"].replace("{company}", contact.get("company", "your company"))
    body = body.replace("{hr_name}", contact.get("name", "Hiring Manager"))
    body = body.replace("{role}", role)
    
    subject = template["subject"].replace("{company}", contact.get("company", "your company"))
    subject = subject.replace("{hr_name}", contact.get("name", "Hiring Manager"))
    subject = subject.replace("{role}", role)
    
    return {
        "subject": subject,
        "body": body,
        "attachments": template.get("attachments", [])
    }

EMAIL_TEMPLATES = load_templates()


# ═══════════════════════════════════════════════════════════════════
# EXCEL PARSER
# ═══════════════════════════════════════════════════════════════════

def load_contacts(filepath=None, start=1, count=None):
    filepath = filepath or config.HR_EXCEL_PATH
    log(f"Loading contacts from: {os.path.basename(filepath)}")
    debug(f"load_contacts(filepath={filepath}, start={start}, count={count})")
    _file_logger.info(f"Loading contacts: file={filepath}, start={start}, count={count}")

    if not os.path.exists(filepath):
        log(f"Excel file not found: {filepath}", "error")
        _file_logger.error(f"Excel file missing: {filepath}")
        return []

    try:
        contacts_data = []
        if filepath.lower().endswith('.csv'):
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f)
                for row in reader:
                    contacts_data.append(row)
        else:
            wb = openpyxl.load_workbook(filepath, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, values_only=True):
                contacts_data.append(row)
            wb.close()
    except Exception as e:
        log(f"Failed to open file: {e}", "error")
        _file_logger.error(f"File open error: {traceback.format_exc()}")
        return []

    contacts = []
    skipped_emails = []
    invalid_rows = 0
    data_started = False
    seen_emails = set()

    try:
        for row_num, row in enumerate(contacts_data, 1):
            # Safe length checks for CSV
            if len(row) < 3:
                invalid_rows += 1
                continue

            if not data_started:
                if row and str(row[0]).strip().startswith("Sr"):
                    data_started = True
                    debug(f"Data header found at row {row_num}")
                    continue
                continue

            if row and len(row) > 2 and row[2] and row[1]:
                email = str(row[2]).strip().replace("\n", "").replace(" ", "")
                company = str(row[1]).strip()
                name = str(row[3]).strip() if len(row) > 3 and row[3] else ""

                # Validate email with regex
                is_valid, reason = validate_email(email)
                if not is_valid:
                    skipped_emails.append((email, reason))
                    debug(f"Row {row_num}: skipped invalid email: {email} ({reason})")
                    continue
                
                email_lower = email.lower()
                if email_lower in seen_emails:
                    skipped_emails.append((email, "Duplicate in list"))
                    debug(f"Row {row_num}: skipped duplicate email: {email}")
                    continue
                seen_emails.add(email_lower)
                
                if not name or name.upper() in ["-", "NA", "N/A", "HR", "TA", "TEAM", "HIRING MANAGER", "TALENT ACQUISITION"]:
                    name = "Hiring Manager"
                else:
                    name = " ".join(name.split())
                    
                if not company or company.upper() in ["-", "NA", "N/A"]:
                    company = "your company"
                else:
                    company = " ".join(company.split())

                try:
                    sr_no = int(row[0]) if row[0] else len(contacts) + 1
                except (ValueError, TypeError):
                    sr_no = len(contacts) + 1

                contacts.append({
                    "sr": sr_no,
                    "company": company,
                    "email": email,
                    "name": name,
                })
            else:
                invalid_rows += 1

    except Exception as e:
        log(f"Error parsing rows: {e}", "error")
        _file_logger.error(f"Data parse error: {traceback.format_exc()}")
        return contacts

    total = len(contacts)
    debug(f"Parsed {total} valid contacts, {len(skipped_emails)} skipped emails, {invalid_rows} empty rows")
    _file_logger.info(f"Parsed: {total} contacts, {len(skipped_emails)} invalid emails, {invalid_rows} empty rows")

    if skipped_emails and DEBUG_MODE:
        log(f"Skipped {len(skipped_emails)} invalid emails (see debug.log)", "warning")
        for em, reason in skipped_emails[:5]:
            debug(f"  Skipped: {em} -> {reason}")

    start_idx = max(0, start - 1)
    if count:
        contacts = contacts[start_idx:start_idx + count]
    else:
        contacts = contacts[start_idx:]

    log(f"Loaded {len(contacts)} contacts (of {total} total)", "success")
    return contacts


# ═══════════════════════════════════════════════════════════════════
# EMAIL GENERATOR (Template-based, no API key needed)
# ═══════════════════════════════════════════════════════════════════

def generate_email(contact, target_role=None):
    """
    Generate a personalized email using smart template rotation.
    Each company gets a deterministic but varied template based on hash.
    """
    debug(f"generate_email: contact={contact}, role={target_role}")
    target_role = target_role or config.TARGET_ROLE
    raw_name = contact.get("name")
    if not raw_name or str(raw_name).strip().upper() in ["", "-", "NA", "N/A", "HR", "TA", "TEAM", "HIRING MANAGER", "TALENT ACQUISITION"]:
        hr_name = "Hiring Manager"
    else:
        hr_name = " ".join(str(raw_name).split())

    raw_company = contact.get("company")
    if not raw_company or str(raw_company).strip().upper() in ["", "-", "NA", "N/A"]:
        company = "your company"
    else:
        company = " ".join(str(raw_company).split())

    if not company or company == "Unknown Company":
        _file_logger.warning(f"generate_email called with missing company: {contact}")

    try:
        hash_val = int(md5(company.encode()).hexdigest(), 16)
        template_idx = hash_val % len(EMAIL_TEMPLATES)
        template = EMAIL_TEMPLATES[template_idx]
        debug(f"Template #{template_idx+1} selected for '{company}' (hash={hash_val % 1000})")

        subject = template["subject"].format(
            hr_name=hr_name, company=company, role=target_role
        )
        body = template["body"].format(
            hr_name=hr_name, company=company, role=target_role
        )

        # Sanity checks on output
        assert len(subject) > 0, "Generated empty subject"
        assert len(body) > 50, f"Generated suspiciously short body ({len(body)} chars)"
        assert "{" not in subject, f"Unresolved placeholder in subject: {subject}"
        debug(f"Generated: subject='{subject[:50]}...', body={len(body)} chars")
        _file_logger.info(f"Email generated for {company}: subject='{subject}'")

        return {
            "subject": subject,
            "body": body,
            "attachments": template.get("attachments", [])
        }

    except KeyError as e:
        _file_logger.error(f"Template placeholder error: {e}\n{traceback.format_exc()}")
        log(f"BUG: Template missing placeholder {e}", "error")
        # Fallback
        return {
            "subject": f"Application for {target_role} - Vinoth Palanivel",
            "body": f"Dear {hr_name},\n\nI am interested in the {target_role} position at {company}. Please find my resume attached.\n\nBest regards,\nVinoth Palanivel\n+91 9500395460\nvinoth322006@gmail.com",
            "attachments": []
        }
    except Exception as e:
        _file_logger.error(f"generate_email crashed: {traceback.format_exc()}")
        log(f"BUG in email generation: {e}", "error")
        return {
            "subject": f"Application for {target_role} - Vinoth Palanivel",
            "body": f"Dear {hr_name},\n\nI am writing regarding opportunities at {company}.\n\nBest regards,\nVinoth Palanivel",
            "attachments": []
        }


# ═══════════════════════════════════════════════════════════════════
# EMAIL SENDER
# ═══════════════════════════════════════════════════════════════════

def send_email(to_email, subject, body, attachments=None, dry_run=False):
    debug(f"send_email(to={to_email}, subject='{subject[:40]}...', dry_run={dry_run})")
    _file_logger.info(f"send_email: to={to_email}, dry_run={dry_run}")

    if dry_run:
        debug("Dry run - skipping actual send")
        return True

    # Pre-send validation
    is_valid, reason = validate_email(to_email)
    if not is_valid:
        log(f"BUG: Invalid recipient email: {to_email} ({reason})", "error")
        _file_logger.error(f"Invalid email rejected: {to_email} -> {reason}")
        return False

    if not config.GMAIL_ADDRESS or not config.GMAIL_APP_PASSWORD:
        log("ERROR: Gmail credentials not set in .env!", "error")
        _file_logger.error("SMTP credentials missing")
        sys.exit(1)

    if attachments is None:
        attachments = [config.DEFAULT_RESUME_PATH]
    elif isinstance(attachments, str):
        attachments = [attachments]

    try:
        debug(f"Building MIME message...")
        msg = MIMEMultipart()
        msg["From"] = f"Vinoth Palanivel <{config.GMAIL_ADDRESS}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        
        domain = config.GMAIL_ADDRESS.split("@")[1] if "@" in config.GMAIL_ADDRESS else "gmail.com"
        msg["Message-ID"] = make_msgid(domain=domain)
        msg["Date"] = formatdate(localtime=True)
        
        msg.attach(MIMEText(body, "plain"))

        for path in attachments:
            if path and os.path.exists(path):
                fsize = os.path.getsize(path)
                debug(f"Attaching: {os.path.basename(path)} ({fsize} bytes)")
                if fsize > 25 * 1024 * 1024:
                    log(f"File too large: {os.path.basename(path)}", "error")
                    continue
                with open(path, "rb") as f:
                    attachment = MIMEBase("application", "octet-stream")
                    attachment.set_payload(f.read())
                    encoders.encode_base64(attachment)
                    attachment.add_header(
                        "Content-Disposition",
                        f"attachment; filename={os.path.basename(path)}",
                    )
                    msg.attach(attachment)

        debug(f"Connecting to {config.SMTP_SERVER}:{config.SMTP_PORT}...")
        with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
            server.send_message(msg)
        _file_logger.info(f"SENT: {to_email} | {subject}")
        debug(f"Email sent successfully to {to_email}")
        return True

    except smtplib.SMTPAuthenticationError as e:
        log("Gmail auth failed! Check your App Password.", "error")
        log("Generate one at: https://myaccount.google.com/apppasswords", "info")
        _file_logger.error(f"SMTP Auth Error: {e}\n{traceback.format_exc()}")
        return False
    except smtplib.SMTPRecipientsRefused as e:
        log(f"Recipient refused: {to_email}", "error")
        _file_logger.error(f"Recipient refused: {to_email}: {e}")
        return False
    except smtplib.SMTPException as e:
        log(f"SMTP error: {e}", "error")
        _file_logger.error(f"SMTP error: {traceback.format_exc()}")
        if "5.4.5" in str(e) or "limit exceeded" in str(e).lower():
            return "RATE_LIMIT"
        return False
    except ConnectionError as e:
        log(f"Network error: {e}", "error")
        _file_logger.error(f"Connection error: {traceback.format_exc()}")
        return False
    except Exception as e:
        log(f"Send failed (unexpected): {e}", "error")
        _file_logger.error(f"Unexpected send error: {traceback.format_exc()}")
        return False


# ═══════════════════════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════════════════════

def log_to_csv(contact, subject, status, resume_used):
    debug(f"log_to_csv: #{contact.get('sr')} {contact.get('company')} -> {status}")
    try:
        file_exists = os.path.exists(config.SEND_LOG_PATH)
        with open(config.SEND_LOG_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow([
                    "Timestamp", "Sr.No", "Company", "Email",
                    "Name", "Subject", "Resume", "Status"
                ])
            writer.writerow([
                datetime.now().isoformat(), contact["sr"], contact["company"],
                contact["email"], contact["name"], subject, resume_used, status,
            ])
    except PermissionError:
        log(f"Cannot write to log file (in use?): {config.SEND_LOG_PATH}", "error")
        _file_logger.error(f"CSV write permission error: {traceback.format_exc()}")
    except Exception as e:
        log(f"CSV logging failed: {e}", "error")
        _file_logger.error(f"CSV logging error: {traceback.format_exc()}")


def get_sent_emails():
    debug(f"get_sent_emails: reading {config.SEND_LOG_PATH}")
    sent = set()
    if not os.path.exists(config.SEND_LOG_PATH):
        debug("No send log file found, returning empty set")
        return sent
    try:
        with open(config.SEND_LOG_PATH, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row_num, row in enumerate(reader, 1):
                if row.get("Status") == "SENT":
                    sent.add(row.get("Email", ""))
        debug(f"Found {len(sent)} previously sent emails")
    except Exception as e:
        log(f"Error reading send log: {e}", "error")
        _file_logger.error(f"Send log read error: {traceback.format_exc()}")
    return sent


def check_bounces():
    """Connect to Gmail via IMAP, find bounce emails, mark as FAILED in CSV, and delete them."""
    if not config.GMAIL_ADDRESS or not config.GMAIL_APP_PASSWORD:
        return
    
    debug("Checking for bounced emails via IMAP...")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
        mail.select("inbox")

        # Search for failure notices
        status, messages = mail.search(None, '(FROM "mailer-daemon@googlemail.com")')
        if status != "OK" or not messages[0]:
            mail.logout()
            return

        bounced_emails = set()
        
        for num in messages[0].split():
            status, data = mail.fetch(num, '(RFC822)')
            if status != "OK":
                continue
                
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            else:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

            failed_rcpt = msg.get("X-Failed-Recipients")
            if failed_rcpt:
                bounced_emails.add(failed_rcpt.strip().lower())
            else:
                match = re.search(r"Your message wasn't delivered to\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", body, re.IGNORECASE)
                if match:
                    bounced_emails.add(match.group(1).strip().lower())

            # Delete the bounce email
            mail.store(num, '+FLAGS', '\\Deleted')

        mail.expunge()
        mail.logout()

        if bounced_emails and os.path.exists(config.SEND_LOG_PATH):
            debug(f"Found {len(bounced_emails)} bounces. Updating log...")
            _file_logger.info(f"Bounces detected: {bounced_emails}")
            rows = []
            with open(config.SEND_LOG_PATH, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                rows.append(header)
                for row in reader:
                    if len(row) > 7 and row[3].strip().lower() in bounced_emails and row[7] == "SENT":
                        row[7] = "FAILED (Bounced)"
                    rows.append(row)
            
            with open(config.SEND_LOG_PATH, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
                
    except Exception as e:
        _file_logger.error(f"Error checking bounces: {traceback.format_exc()}")


# ═══════════════════════════════════════════════════════════════════
# MAIN ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════

def run_automail(args):
    global _start_time
    _start_time = time.time()
    _file_logger.info(f"run_automail called with: {vars(args)}")
    debug(f"Args: {vars(args)}")
    banner()

    # ── List resumes ──
    if args.list_resumes:
        print(f"\n  {Colors.BOLD}Available Resumes:{Colors.RESET}\n")
        for key, info in config.AVAILABLE_RESUMES.items():
            exists = "+" if os.path.exists(info["path"]) else "x"
            color = Colors.GREEN if os.path.exists(info["path"]) else Colors.RED
            print(f"    {color}{exists}{Colors.RESET}  {Colors.BOLD}{key}{Colors.RESET}")
            print(f"       {Colors.DIM}{info['label']} -> {info['filename']}{Colors.RESET}")
        print()
        return

    # ── Resolve resume ──
    resume_path = config.DEFAULT_RESUME_PATH
    resume_label = config.DEFAULT_RESUME
    if args.resume:
        if args.resume in config.AVAILABLE_RESUMES:
            resume_info = config.AVAILABLE_RESUMES[args.resume]
            resume_path = resume_info["path"]
            resume_label = resume_info["filename"]
        else:
            log(f"Unknown resume key: {args.resume}. Use --list-resumes", "error")
            return

    if not os.path.exists(resume_path):
        log(f"Resume file not found: {resume_path}", "error")
        return

    # ── Load contacts ──
    contacts = load_contacts(start=args.start, count=args.count)
    if not contacts:
        log("No contacts loaded!", "error")
        return

    # ── Skip already sent ──
    if not args.no_skip:
        sent_emails = get_sent_emails()
        original_count = len(contacts)
        contacts = [c for c in contacts if c["email"] not in sent_emails]
        skipped = original_count - len(contacts)
        if skipped > 0:
            log(f"Skipped {skipped} already-sent contacts", "warning")

    if not contacts:
        log("All contacts in this range already emailed!", "warning")
        return

    # ── Config ──
    target_role = args.role or config.TARGET_ROLE
    mode = f"{Colors.YELLOW}DRY RUN{Colors.RESET}" if args.dry_run else f"{Colors.GREEN}LIVE{Colors.RESET}"

    print(f"\n  {Colors.BOLD}--- Batch Configuration ---{Colors.RESET}")
    print(f"  Mode:       {mode}")
    print(f"  Contacts:   {Colors.BOLD}{len(contacts)}{Colors.RESET}")
    print(f"  Range:      #{contacts[0]['sr']} -> #{contacts[-1]['sr']}")
    print(f"  Resume:     {Colors.CYAN}{resume_label}{Colors.RESET}")
    print(f"  Target:     {Colors.CYAN}{target_role}{Colors.RESET}")
    print(f"  Templates:  {Colors.CYAN}{len(EMAIL_TEMPLATES)} AI-generated variants{Colors.RESET}")
    print(f"  Delay:      {config.DELAY_BETWEEN_EMAILS}s between emails")
    print()

    if not args.dry_run and not args.test and not args.yes:
        confirm = input(f"  {Colors.YELLOW}Ready to send {len(contacts)} emails? (y/N): {Colors.RESET}")
        if confirm.lower() != "y":
            log("Aborted by user", "warning")
            return

    # ── Process ──
    success_count = 0
    fail_count = 0

    for i, contact in enumerate(contacts, 1):
        print(f"\n  {Colors.BOLD}[{i}/{len(contacts)}]{Colors.RESET} {Colors.CYAN}{contact['company']}{Colors.RESET}")
        print(f"           -> {contact['email']} ({contact['name']})")

        # Generate email
        log(f"Generating email for {contact['company']}...", "ai")
        email_data = generate_email(contact, target_role)
        subject = email_data["subject"]
        body = email_data["body"]

        # Preview
        print(f"\n    {Colors.DIM}{'=' * 50}{Colors.RESET}")
        print(f"    {Colors.BOLD}Subject:{Colors.RESET} {subject}")
        print(f"    {Colors.DIM}{'-' * 50}{Colors.RESET}")
        for line in body.split("\n")[:6]:
            print(f"    {Colors.DIM}{line}{Colors.RESET}")
        if len(body.split("\n")) > 6:
            print(f"    {Colors.DIM}... [{len(body.split(chr(10)))} lines total]{Colors.RESET}")
        print(f"    {Colors.DIM}{'=' * 50}{Colors.RESET}")

        # Send or dry-run
        if args.dry_run:
            log("DRY RUN - email not sent", "warning")
            status = "DRY_RUN"
            success_count += 1
        elif args.test:
            log(f"TEST MODE - sending to {config.GMAIL_ADDRESS}", "warning")
            sent = send_email(config.GMAIL_ADDRESS, subject, body, resume_path)
            status = "SENT" if sent else "FAILED"
            if sent:
                success_count += 1
                log(f"Test email sent to {config.GMAIL_ADDRESS}!", "send")
            else:
                fail_count += 1
            log_to_csv(contact, subject, status, resume_label)
            break
        else:
            sent = send_email(contact["email"], subject, body, resume_path)
            status = "SENT" if sent else "FAILED"
            if sent:
                success_count += 1
                log(f"Sent to {contact['email']}", "send")
            else:
                fail_count += 1
                log(f"Failed for {contact['email']}", "error")

        log_to_csv(contact, subject, status, resume_label)

        if i < len(contacts) and not args.dry_run:
            log(f"Waiting {config.DELAY_BETWEEN_EMAILS}s (rate limit)...", "wait")
            time.sleep(config.DELAY_BETWEEN_EMAILS)

    # ── Summary ──
    elapsed = time.time() - _start_time
    print(f"\n  {Colors.BOLD}--- Results ---{Colors.RESET}")
    print(f"  {Colors.GREEN}+ Successful: {success_count}{Colors.RESET}")
    if fail_count > 0:
        print(f"  {Colors.RED}x Failed:     {fail_count}{Colors.RESET}")
    print(f"  Time:   {elapsed:.1f}s")
    print(f"  Log:    {config.SEND_LOG_PATH}")
    print(f"  Debug:  {_log_path}")
    _file_logger.info(f"Batch complete: {success_count} sent, {fail_count} failed, {elapsed:.1f}s")
    print()


# ═══════════════════════════════════════════════════════════════════
# CLI INTERFACE
# ═══════════════════════════════════════════════════════════════════

def main():
    global DEBUG_MODE

    parser = argparse.ArgumentParser(
        description="AUTOMAIL - AI-Powered Cold Email Automation (No API Key Needed)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python automail.py --dry-run --count 3           Preview 3 emails
  python automail.py --test                         Send test to yourself
  python automail.py --start 1 --count 50           Send to contacts 1-50
  python automail.py --resume vinoth_skct --count 10
  python automail.py --role "Data Scientist Intern" --count 20
  python automail.py --health-check                 Run diagnostics
  python automail.py --debug --dry-run --count 2    Verbose debug output
        """
    )
    parser.add_argument("--dry-run", action="store_true", help="Preview mode (no sending)")
    parser.add_argument("--test", action="store_true", help="Send test to yourself")
    parser.add_argument("--start", type=int, default=1, help="Starting contact # (default: 1)")
    parser.add_argument("--count", type=int, default=None, help="Number of contacts")
    parser.add_argument("--resume", type=str, default=None, help="Resume key (see --list-resumes)")
    parser.add_argument("--list-resumes", action="store_true", help="List resume files")
    parser.add_argument("--role", type=str, default=None, help="Target role")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip confirmation")
    parser.add_argument("--no-skip", action="store_true", help="Don't skip already-sent")
    parser.add_argument("--debug", action="store_true", help="Enable verbose debug output")
    parser.add_argument("--health-check", action="store_true", help="Run system diagnostics")

    args = parser.parse_args()

    # Enable debug mode
    if args.debug:
        DEBUG_MODE = True
        debug("Debug mode ENABLED")
        debug(f"Python {sys.version}")
        debug(f"Working dir: {os.getcwd()}")
        debug(f"Config base: {config.BASE_DIR}")

    # Health check mode
    if args.health_check:
        banner()
        health_check()
        return

    if not any([args.dry_run, args.test, args.list_resumes, args.count, args.start != 1]):
        log("No flags specified. Running DRY-RUN with 3 contacts.", "warning")
        args.dry_run = True
        args.count = 3

    try:
        run_automail(args)
    except KeyboardInterrupt:
        print(f"\n  {Colors.YELLOW}Interrupted by user{Colors.RESET}")
        _file_logger.warning("Session interrupted by user (Ctrl+C)")
    except Exception as e:
        log(f"FATAL ERROR: {e}", "error")
        _file_logger.critical(f"Unhandled exception:\n{traceback.format_exc()}")
        print(f"\n  {Colors.RED}Full traceback saved to: {_log_path}{Colors.RESET}")
        if DEBUG_MODE:
            traceback.print_exc()


if __name__ == "__main__":
    main()
