import pandas as pd
import jinja2 as jin
import smtplib as smt
from email.message import EmailMessage

#Obtain template name from user
#Obtain the Excel Sheets from user
#For each entry in the excel sheet
#   Obtain all valid info (replace null if no info, inform user)
#   Fill in the template
#   Send it to the email as a message

#Processes the mailmerge
def process_mail_merge(sheet_path, template_path, from_email, password):
    excelDf = pd.read_excel(sheet_path, engine="openpyxl")

    for i, row in excelDf.iterrows():
        name = row.get("Name", "").strip() if pd.notna(row.get("Name")) else ""
        email = row.get("Email Address", "").strip() if pd.notna(row.get("Email Address")) else ""

        if not name or not email:
            print(f"[!] Row {i + 2}: Missing Name or Email — Skipping.")
            continue

        filledTemplate = fillInTemplate(template_path, name)
        filledSubject = fillInSubject(name)
        sendEmail(to_email=email, subject=filledSubject, body=filledTemplate, from_email=from_email, password=password)
        print(f"[✓] Sent email to {name} at {email}")

#Fills in the template with the corresponding labels
def fillInTemplate(templatePath, name):
    with open(templatePath, "r", encoding="utf-8") as file:
        template = jin.Template(file.read())
    return template.render(Name=name)

#Fills in a subject line
def fillInSubject(name):
    return f"Hello There {name}!"

#Sends the email to all valid entries in the database
def sendEmail(to_email, subject, body, from_email, password):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(body)

    with smt.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(from_email, password)
        smtp.send_message(msg)

