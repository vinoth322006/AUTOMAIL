import config
from automail import EMAIL_TEMPLATES

for i, template in enumerate(EMAIL_TEMPLATES):
    subject = template["subject"].format(role="AI/ML Engineer Intern", company="Test Company", hr_name="HR Manager")
    print(f"Template {i+1}:")
    print(f"Subject: {subject}")
    print("-" * 40)
