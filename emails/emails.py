import os
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Emails(object):
    def __init__(self):
        self.smtp_server = os.getenv('SMTP_SERVER')
        self.port = int(os.getenv('SMTP_PORT'))
        self.sender_email = os.getenv('EMAIL_SENDER')
        self.subject = os.getenv('EMAIL_SUBJECT')
        with open(os.path.join(os.path.dirname(__file__), 'logo.png'), 'rb') as file:
            self.logo_img = file.read()
        with open(os.path.join(os.path.dirname(__file__), 'email_template.html'), 'r', encoding='UTF-8') as file:
            self.html = file.read()

    def extract_name(self, email):
        name = email.split('@')[0]
        return name.split('.')[0].capitalize()

    def send_email(self, receiver, info):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = receiver
        html_part = MIMEMultipart("related")
        first_name = self.extract_name(receiver)
        self.html = self.html.replace('${FIRST_NAME}', first_name)
        sender_name = self.extract_name(receiver)
        self.html = self.html.replace('${SENDER_NAME}', sender_name)
        self.html = self.html.replace('${TABLE}', info)
        html_part.attach(MIMEText(self.html, "html"))
        logo_image = MIMEImage(self.logo_img)
        logo_image.add_header('Content-ID', '<logo>')
        html_part.attach(logo_image)
        message.attach(html_part)

        print(f'Send email to {receiver}')
        with smtplib.SMTP(self.smtp_server, self.port) as server:
            server.sendmail(from_addr=self.sender_email, to_addrs=[receiver], msg=message.as_string())
