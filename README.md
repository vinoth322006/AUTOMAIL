# ⚡ AUTOMAIL: High-Performance AI Cold Email Outreach

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Build: Production Ready](https://img.shields.io/badge/Build-Production--Ready-brightgreen.svg)](#)

**Automail** is a professional, production-grade automated outreach system designed to secure internships and jobs. It combines AI-driven template rotation, a sleek web dashboard, and smart delivery protection to ensure your cold emails land in the inbox, not the spam folder.

---

## ✨ Key Features

- **🌐 Smart Web Dashboard**: A fully interactive GUI for managing contacts, settings, and live campaign monitoring.
- **🤖 AI Template Rotation**: Automatically rotates through multiple professional, fact-dense email templates to bypass spam filters and maintain high engagement.
- **🛡️ Delivery Protection**:
    - **Bounce Detection**: Automatically checks your Gmail via IMAP for bounces and marks failed domains.
    - **Rate Limiting**: Configurable delays (4-10s) and daily quotas to stay within Gmail safety limits.
    - **Deduplication**: Never sends the same email twice to the same recipient.
- **📁 Advanced Asset Management**: Support for multiple resumes and template-specific attachments.
- **📊 Real-time Monitoring**: Live progress bars, health checks, and detailed audit logs.
- **📂 Flexible Data Sources**: Seamlessly import contacts from `.xlsx` (Excel) or `.csv` files.

---

## 🚀 Quick Start

### 1. Prerequisites
Ensure you have **Python 3.8+** installed.

### 2. Installation
Clone the repository and install the required dependencies:
```bash
git clone https://github.com/your-username/AUTOMAIL.git
cd AUTOMAIL
pip install -r requirements.txt
```

### 3. Launch the Dashboard
Start the local server:
```bash
python server.py
```
Open your browser and navigate to: **[http://localhost:8000](http://localhost:8000)**

---

## ⚙️ Configuration

Automail uses a `.env` file for secure credential management.

1.  Copy `.env.example` to `.env`.
2.  Add your **Gmail Address**.
3.  Add your **Gmail App Password**.
    > [!IMPORTANT]
    > You **must** use a 16-character App Password, NOT your regular Google password.
    > 1. Go to [Google App Passwords](https://myaccount.google.com/apppasswords).
    > 2. Create a new app named "Automail".
    > 3. Copy the 16-character code into your `.env`.

### `config.py` Overview
You can further customize delivery logic in `config.py`:
- `DELAY_BETWEEN_EMAILS`: Seconds to wait between sends (default: 4s).
- `MAX_DAILY_EMAILS`: Safety ceiling for your Gmail account.
- `TARGET_ROLE`: Your default job/internship target.

---

## 🖥️ Dashboard Overview

- **Dashboard Tab**: Monitor sending progress, view daily stats, and manage SMTP health.
- **Contacts Tab**: Upload new contact lists, add single leads, or edit existing records.
- **Templates Tab**: Customize your AI-generated templates and associate specific attachments with them.
- **Logs Tab**: View a live feed of sent emails, failures, and debug information.

---

## 🛡️ Best Practices for Success

1.  **Warm-up Your Account**: Start by sending 20-30 emails per day and gradually increase to 100+.
2.  **Verify Contacts**: Ensure your Excel sheet has the correct format: `Sr.No`, `Company Name`, `Email Address`, `Recipient Name`.
3.  **Template Variety**: Use the dashboard to add at least 3-5 different templates to maximize the benefits of the rotation system.
4.  **Monitor Bounces**: If your bounce rate exceeds 5%, stop the campaign and verify your lead list.

---

## 🛠️ Tech Stack

- **Backend**: Python 3.8+, `http.server`, `smtplib`, `imaplib`
- **Frontend**: Vanilla HTML5, CSS3 (Modern Glassmorphism Design), JavaScript (ES6)
- **Data**: Excel (`openpyxl`), CSV, JSON
- **AI Integration**: Pre-generated Gemini/Proxima templates

---

## 📜 License

Distributed under the MIT License. See `LICENSE` for more information.

---

**Developed with ❤️ for high-performance job seekers.**
