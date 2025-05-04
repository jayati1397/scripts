#!/usr/bin/env python3
import pandas as pd
import smtplib
import ssl
import mimetypes
from email.message import EmailMessage
import sys
import os
import time
import config

RESUME_FILES = {
    "python": "resume/Jayati_Resume_Python.pdf",
    "devops": "resume/Jayati_Resume_DevOps.pdf",
    "java":   "resume/Jayati_Resume_Java.pdf"
}

def send_email(sender_email, sender_password, recipient, subject, body, attachment_path=None):
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path and os.path.isfile(attachment_path):
        ctype, encoding = mimetypes.guess_type(attachment_path)
        if ctype is None or encoding:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(attachment_path, "rb") as fp:
            msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, sender_password)
        server.send_message(msg)

def main():
    if len(sys.argv) != 3:
        print("Usage: python send_personalized_email.py recipients.xlsx body_template.txt")
        sys.exit(1)

    excel_path = sys.argv[1]
    template_path = sys.argv[2]

    with open(template_path, encoding='utf-8') as f:
        template = f.read()

    df = pd.read_excel(excel_path)

    sender_email = config.GMAIL_ADDRESS
    sender_password = config.GMAIL_APP_PASSWORD
    static_subject = "Interest in Software Engineer roles at "

    for i, row in df.iterrows():
        to_addr = row.get("recipient_email", "").strip()
        name = row.get("name", "").strip()
        company = row.get("company", "").strip()
        resume_key = str(row.get("resume", "")).lower().strip()
        subject = static_subject + row.get("company", "").strip()
        roles_text = row.get("roles", "")

        if not to_addr or not name or not company:
            print(f"Row {i+1}: missing required fields, skipping")
            continue

        roles_list = ""
        if roles_text:
            items = [line.strip() for line in roles_text.splitlines() if line.strip()]
            if items:
                roles_list = "\nSome open roles on careers page:\n"
                roles_list += "\n".join(f"* {item}" for item in items)

        attachment_file = RESUME_FILES.get(resume_key)

        body = template.format(name=name, company=company, roles_list=roles_list)

        try:
            send_email(sender_email, sender_password, to_addr, subject, body, attachment_file)
            print(f"[{i+1}] Sent to {to_addr} with attachment {attachment_file}")
        except Exception as e:
            print(f"[{i+1}] Failed to send to {to_addr}: {e}")

        time.sleep(10)

if __name__ == "__main__":
    main()
