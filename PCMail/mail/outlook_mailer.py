# mail/outlook_mailer.py

import win32com.client

from PCMail.config.recipients import TO_RECIPIENTS
from statics.data import STRUCTURED_INVESTMENTS_EMAIL

def send_mail(subject, html_body, attachments):

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = TO_RECIPIENTS
    mail.CC = STRUCTURED_INVESTMENTS_EMAIL
    mail.Subject = subject

    mail.Display()
    signature = mail.HTMLBody
    mail.HTMLBody = html_body + signature

    for f in attachments:
        mail.Attachments.Add(f)
