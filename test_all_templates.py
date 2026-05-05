"""Send all 6 email templates to yourself for review."""
import time
import sys
import io

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import config
from automail import EMAIL_TEMPLATES, send_email, log, Colors

FAKE_COMPANIES = [
    # Each company maps to a different template via md5 hash
    ("Google", "Sundar Pichai"),
    ("Microsoft", "Satya Nadella"),
    ("Amazon", "Andy Jassy"),
    ("Tesla", "Elon Musk"),
    ("Meta", "Mark Zuckerberg"),
    ("Apple", "Tim Cook"),
]

role = config.TARGET_ROLE
to = config.GMAIL_ADDRESS

print(f"\n  {Colors.BOLD}Sending all 6 templates to: {to}{Colors.RESET}\n")

# Figure out which company maps to which template
from hashlib import md5
company_template_map = []
used_templates = set()

for company, name in FAKE_COMPANIES:
    idx = int(md5(company.encode()).hexdigest(), 16) % len(EMAIL_TEMPLATES)
    if idx not in used_templates:
        used_templates.add(idx)
        company_template_map.append((company, name, idx))

# Fill in any missing templates with custom names
for i in range(len(EMAIL_TEMPLATES)):
    if i not in used_templates:
        company_template_map.append((f"TestCorp{i}", "HR Manager", i))
        used_templates.add(i)

# Sort by template index
company_template_map.sort(key=lambda x: x[2])

sent = 0
for company, hr_name, idx in company_template_map:
    template = EMAIL_TEMPLATES[idx]
    subject = f"[TEMPLATE {idx+1}/6] " + template["subject"].format(
        hr_name=hr_name, company=company, role=role
    )
    body = template["body"].format(hr_name=hr_name, company=company, role=role)

    print(f"  [{idx+1}/6] Template #{idx+1} — {Colors.CYAN}{subject[:60]}...{Colors.RESET}")

    ok = send_email(to, subject, body, config.DEFAULT_RESUME_PATH)
    if ok:
        sent += 1
        log(f"Sent template {idx+1}/6", "send")
    else:
        log(f"Failed template {idx+1}/6", "error")

    if idx < len(EMAIL_TEMPLATES) - 1:
        time.sleep(5)

print(f"\n  {Colors.GREEN}{Colors.BOLD}Done! {sent}/6 templates sent to {to}{Colors.RESET}")
print(f"  Check your inbox!\n")
