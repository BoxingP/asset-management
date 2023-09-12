import datetime
import os
import re
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

from utils.email_logger import EmailSendingLogger
from utils.logger import Logger


class Emails(object):
    def __init__(self, name):
        self.smtp_server = os.getenv('SMTP_SERVER')
        self.port = int(os.getenv('SMTP_PORT'))
        self.sender_email = os.getenv(f'{name.upper()}_EMAIL_SENDER')
        self.subject = os.getenv(f'{name.upper()}_EMAIL_SUBJECT')
        with open(os.path.join(os.path.dirname(__file__), 'signature.png'), 'rb') as file:
            self.signature_img = file.read()
        with open(os.path.join(os.path.dirname(__file__), f'{name}_template.html'), 'r', encoding='UTF-8') as file:
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

    def send_email(self, sender, to: list, email_content, cc: list = None, bcc: list = None, record_sent: bool = False):
        receivers = [item for sublist in (to, cc, bcc) if sublist is not None for item in sublist]
        Logger().info(msg=f"Will send email to {', '.join(receivers)}")
        logger = EmailSendingLogger()
        try:
            with smtplib.SMTP(self.smtp_server, self.port) as server:
                send_errs = server.sendmail(from_addr=sender, to_addrs=receivers, msg=email_content.as_string())
            if not send_errs:
                Logger().info(msg=f"Successfully sent email to {', '.join(receivers)}")
                if record_sent:
                    for recipient in to:
                        logger.log_email_sent(recipient=recipient, subject=email_content["Subject"], success=True)
            else:
                for key, value in send_errs.items():
                    if key == os.getenv('RETURN_EMAIL_CC'):
                        continue
                    code, message = value
                    error_message = message.decode('utf-8')
                    Logger().error(msg=f'Failed to send email to {key}: {code} - {error_message}')
                    if record_sent:
                        logger.log_email_sent(recipient=key, subject=email_content["Subject"], success=False,
                                              error_code=code, error_message=error_message)

        except smtplib.SMTPRecipientsRefused as e:
            for recipient, (code, message) in e.recipients.items():
                error_message = message.decode('utf-8')
                Logger().error(msg=f'Failed to send email to {recipient}: {code} - {error_message}')
                if record_sent:
                    logger.log_email_sent(recipient=recipient, subject=email_content["Subject"], success=False,
                                          error_code=code, error_message=error_message)

    def send_asset_email(self, receiver, info):
        message = MIMEMultipart("alternative")
        subject = self.subject.replace('YEAR', str(datetime.datetime.now().year))
        message["Subject"] = subject
        message["From"] = self.sender_email
        message["To"] = receiver
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${RECEIVER}', self.extract_name(receiver, is_firstname=True))
        self.html = self.html.replace('${ASSET_URL}', os.getenv('ASSET_EMAIL_ASSET_URL'))
        self.html = self.html.replace('${IT_SUPPORT_EMAIL}', os.getenv('ASSET_EMAIL_IT_SUPPORT_MAILBOX'))
        self.html = self.html.replace('${EMAIL_SUBJECT}', subject)
        self.html = self.html.replace('${TABLE}', info)
        self.html = self.html.replace('${SENDER}', self.extract_name(self.sender_email))
        self.html = self.html.replace('${SENDER_TITLE}', os.getenv('ASSET_EMAIL_SENDER_TITLE'))
        self.html = self.html.replace('${SENDER_ADDRESS}',
                                      '<br>'.join(os.getenv('ASSET_EMAIL_SENDER_ADDRESS').split(';')))
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)

        self.send_email(sender=self.sender_email, to=[receiver], email_content=message)

    def send_return_email(self, emp_id, name, email, date, info):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject.replace('NAME', str(name)).replace('ID', str(emp_id))
        message["From"] = self.sender_email
        message["To"] = email
        message["Cc"] = os.getenv('RETURN_EMAIL_CC')
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${RECEIVER}', name)
        self.html = self.html.replace('${DATE}', date.strftime('%Y年%m月%d日'))
        self.html = self.html.replace('${IT_CONTACT}', self.generate_it_contact())
        self.html = self.html.replace('${TABLE}', info)
        self.html = self.html.replace('${IT_SUPPORT_EMAIL}', os.getenv('RETURN_EMAIL_IT_SUPPORT_MAILBOX'))
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)

        self.send_email(sender=self.sender_email, to=[email], cc=[os.getenv('RETURN_EMAIL_CC')], email_content=message,
                        record_sent=True)

    def send_return_error_email(self, info, excel_attachment=None):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = self.sender_email
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${TABLE}', info)
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)
        if excel_attachment is not None:
            message.attach(excel_attachment)

        self.send_email(sender=self.sender_email, to=[self.sender_email], email_content=message)

    def send_return_summary_email(self, info, excel_attachment=None):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = self.sender_email
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${TABLE}', info)
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)
        if excel_attachment is not None:
            message.attach(excel_attachment)

        self.send_email(sender=self.sender_email, to=[self.sender_email], email_content=message)

    def generate_it_contact(self):
        df = pd.read_csv(os.path.join(os.path.dirname(__file__), 'it_contact.csv'), sep=',')
        df['驻场工程师'] = df['驻场工程师'].str.replace(',', '<br>')
        df['联系电话'] = df['联系电话'].str.replace(',', '<br>')
        df['联系邮箱'] = df['联系邮箱'].str.replace(',', '<br>')
        return df.to_html(index=False, escape=False)

    def extract_year_quarter(self):
        pattern = r'(\d{4})(Q\d)'
        match = re.search(pattern, self.subject)
        if match:
            year = match.group(1)
            quarter = match.group(2)
        else:
            current_date = datetime.datetime.now()
            year = str(current_date.year)
            quarter = f'Q{str((current_date.month - 1) // 3 + 1)}'
        return year, quarter

    def send_inventory_email(self, name, email, info):
        year, quarter = self.extract_year_quarter()
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject.replace('YEAR', year).replace('QUARTER', quarter)
        message["From"] = self.sender_email
        message["To"] = email
        message["Cc"] = os.getenv('QUARTERLY_ASSET_EMAIL_CC')
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${RECEIVER}', name)
        self.html = self.html.replace('${YEAR}', year)
        self.html = self.html.replace('${QUARTER}', quarter)
        self.html = self.html.replace('${TABLE}', info)
        self.html = self.html.replace('${ASSET_URL}', os.getenv('QUARTERLY_ASSET_EMAIL_ASSET_URL'))
        self.html = self.html.replace('${IT_SUPPORT_EMAIL}', os.getenv('QUARTERLY_ASSET_EMAIL_IT_SUPPORT_MAILBOX'))
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        with open(os.path.join(os.path.dirname(__file__), 'confirm.png'), 'rb') as file:
            confirm_img = file.read()
        confirm_image = MIMEImage(confirm_img)
        confirm_image.add_header('Content-ID', '<confirm>')
        html_part.attach(confirm_image)
        with open(os.path.join(os.path.dirname(__file__), 'feedback.png'), 'rb') as file:
            feedback_img = file.read()
        feedback_image = MIMEImage(feedback_img)
        feedback_image.add_header('Content-ID', '<feedback>')
        html_part.attach(feedback_image)
        message.attach(html_part)

        self.send_email(sender=self.sender_email, to=[email], cc=[os.getenv('QUARTERLY_ASSET_EMAIL_CC')],
                        email_content=message, record_sent=True)

    def send_inventory_error_email(self, info, excel_attachment=None):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = self.sender_email
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${TABLE}', info)
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)
        if excel_attachment is not None:
            message.attach(excel_attachment)

        self.send_email(sender=self.sender_email, to=[self.sender_email], email_content=message)

    def send_inventory_summary_email(self, info, excel_attachment=None):
        message = MIMEMultipart("alternative")
        message["Subject"] = self.subject
        message["From"] = self.sender_email
        message["To"] = self.sender_email
        html_part = MIMEMultipart("related")
        self.html = self.html.replace('${TABLE}', info)
        html_part.attach(MIMEText(self.html, "html"))
        signature_image = MIMEImage(self.signature_img)
        signature_image.add_header('Content-ID', '<signature>')
        html_part.attach(signature_image)
        message.attach(html_part)
        if excel_attachment is not None:
            message.attach(excel_attachment)

        self.send_email(sender=self.sender_email, to=[self.sender_email], email_content=message)
