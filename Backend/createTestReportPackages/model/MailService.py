import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from createTestReportPackages.parser import CONFIG


class MailService:
    def __init__(self):
        try:
            if self.MAIL_SERVER:
                pass
        except Exception as e:
            self.initializeMail()

    def initializeMail(self):
        self.MAIL_SERVER = smtplib.SMTP(CONFIG["mail_server"], 587)
        self.MAIL_SERVER.starttls()
        print("Logging in")
        self.MAIL_SERVER.login(CONFIG["from"], CONFIG["password"])
        print("Logging Successful")

    def create_message(self,mail_data):
        try:
            print("creating message")
            msg = MIMEMultipart()
            msg['From'] = CONFIG["from"]
            msg['To'] = mail_data["to"]
            msg['Subject'] = mail_data["Subject"]
            body = mail_data["body"]
            msg.attach(MIMEText(body, 'html'))
            print("message created")
            return msg
        except Exception as e:
            print(e)

    def send_mail(self, mail_data):
        try:
            if False:
                msg = self.create_message(mail_data)
                print("trying to send mail ....")
                text = msg.as_string()
                self.MAIL_SERVER.sendmail(CONFIG["from"], mail_data["to"].split(","), text)
                self.MAIL_SERVER.quit()
                print("mail sent")
        except Exception as e:
            print(e)
