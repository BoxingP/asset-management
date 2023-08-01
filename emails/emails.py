import datetime
import os
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from utils.logger import Logger


class Emails(object):
    def __init__(self):
        self.smtp_server = os.getenv('SMTP_SERVER')
        self.port = int(os.getenv('SMTP_PORT'))
        self.sender_email = os.getenv('EMAIL_SENDER')
        self.subject = os.getenv('EMAIL_SUBJECT').replace('YEAR', str(datetime.datetime.now().year))
        with open(os.path.join(os.path.dirname(__file__), 'signature.png'), 'rb') as file:
            self.signature_img = file.read()
        with open(os.path.join(os.path.dirname(__file__), 'email_template.html'), 'r', encoding='UTF-8') as file:
            self.html = file.read()

    def extract_name(self, email, is_firstname=False):
        username, domain = email.split('@')
        name_parts = username.split('.')
        if len(name_parts) >= 2:
            first_name = '.'.join(name_parts[:-1]).capitalize()
            last_name = name_parts[-1].capitalize()
            full_name = f'{first_name} {last_name}'
        else:
            full_name = first_name = username.capitalize()
        if is_firstname:
            return first_name
        else:
            return full_name

    def send_email(self, receiver, info):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = receiver
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${RECEIVER}', self.extract_name(receiver, is_firstname=True))
        self.html = self.html.replace('${ASSET_URL}', os.getenv('EMAIL_ASSET_URL'))
        self.html = self.html.replace('${IT_SUPPORT_EMAIL}', os.getenv('EMAIL_IT_SUPPORT_MAILBOX'))
        self.html = self.html.replace('${EMAIL_SUBJECT}', self.subject)
        self.html = self.html.replace('${TABLE}', info)
        self.html = self.html.replace('${SENDER}', self.extract_name(os.getenv('EMAIL_SENDER')))
        self.html = self.html.replace('${SENDER_TITLE}', os.getenv('EMAIL_SENDER_TITLE'))
        self.html = self.html.replace('${SENDER_ADDRESS}', '<br>'.join(os.getenv('EMAIL_SENDER_ADDRESS').split(';')))
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)

        Logger().info(msg=f'Will send email to {receiver}')
        try:
            with smtplib.SMTP(self.smtp_server, self.port) as server:
                server.sendmail(from_addr=self.sender_email, to_addrs=[receiver], msg=message.as_string())
            Logger().info(msg=f'Successfully sent email to {receiver}')
        except smtplib.SMTPRecipientsRefused as e:
            for recipient, (code, message) in e.recipients.items():
                Logger().error(msg=f'Failed to send email to {recipient}: {code} - {message}')
