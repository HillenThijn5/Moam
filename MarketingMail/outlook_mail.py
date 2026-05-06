import html

import win32com.client as win32

from MarketingMail.config import TO_LIST, CC_LIST

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001F"

class OutlookMailer:
    def __init__(self):
        self.outlook = None

    def __enter__(self):
        try:
            self.outlook = win32.GetObject(Class="Outlook.Application")
        except Exception:
            self.outlook = win32.Dispatch("Outlook.Application")
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def create_mail_with_inline_image(self, subject: str, body_text: str, image_path: str):
        mail = self.outlook.CreateItem(0)

        mail.Subject = subject
        if TO_LIST:
            mail.To = TO_LIST
        if CC_LIST:
            mail.CC = CC_LIST

        # attach image + set CID
        att = mail.Attachments.Add(image_path)
        cid = "rangeimg"
        pa = att.PropertyAccessor
        pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
        pa.SetProperty(PR_ATTACH_CONTENT_LOCATION, cid)

        # HTML body (image inline + tekst eronder)
        safe_text = html.escape(body_text).replace("\r\n", "<br>")
        mail.HTMLBody = f"""
        <html>
          <body>
            <img src="cid:{cid}"><br><br>
            {safe_text}
          </body>
        </html>
        """

        mail.Display()
        return mail



