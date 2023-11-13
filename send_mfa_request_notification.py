import errno
import os
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

from databases.emp_collect_database import EMPCollectDatabase
from databases.emp_info_database import EMPInfoDatabase
from emails.emails import Emails
from utils.email_logger import EmailSendingLogger


def generate_summary_data():
    df = pd.DataFrame(EmailSendingLogger().get_log_data())
    df['success'] = df['success'].replace({True: 'Y', False: 'N'})
    df = df[['time', 'recipient', 'subject', 'success', 'error_message']]
    df.rename(columns={'time': '发送时间', 'recipient': '收件邮箱', 'subject': '邮件标题', 'success': '是否发送成功',
                       'error_message': '详细信息'}, inplace=True)
    soup = BeautifulSoup(df.to_html(index=False), 'html.parser')
    rows = soup.find_all('tbody')[0].find_all('tr')
    for row in rows:
        success_cell = row.find_all('td')[3]
        if success_cell.get_text() == 'N':
            row['class'] = 'failed'
    return soup.prettify()


def attach_excel(excel_file):
    with open(excel_file, 'rb') as attachment:
        execl = MIMEBase("application", "octet-stream")
        execl.set_payload(attachment.read())
    encoders.encode_base64(execl)
    execl.add_header('Content-Disposition', 'attachment', filename=Header(excel_file.name, 'utf-8').encode())
    return execl


def generate_issue_data(dataframe, validate_result, key_column):
    result_dict = {}
    dataframe[key_column] = dataframe[key_column].str.lower()
    for entry in validate_result:
        index = entry['index'].lower()
        reason = ', '.join(entry['reason'])

        filtered_df = dataframe[dataframe[key_column] == index].copy()
        if not filtered_df.empty:
            filtered_df.loc[:, '原因'] = reason
            result_dict[index] = filtered_df
    result_df = pd.concat(result_dict.values(), ignore_index=True)
    result_df.replace({None: '', pd.NA: '', float('nan'): ''}, inplace=True)
    return result_df


def validate_band(email, emp_info_database, emp_collect_database):
    if emp_info_database.is_china_vip(email) or emp_collect_database.is_high_band(email):
        return {'index': email, 'reason': '员工级别高'}


def validate_info(dataframe, email_column):
    validate_result = []
    mp_info_database = EMPInfoDatabase()
    emp_collect_database = EMPCollectDatabase()
    for email in dataframe[email_column]:
        if email is not None:
            validate_result.append(validate_band(email, mp_info_database, emp_collect_database))
    validate_result = list(filter(lambda item: item is not None, validate_result))
    result = {}
    for entry in validate_result:
        index = entry['index']
        reason = entry['reason']
        if index in result:
            result[index]['reason'].append(reason)
        else:
            result[index] = {'index': index, 'reason': [reason]}
    result_list = list(result.values())
    return result_list


def get_visible_sheet_name(excel_file):
    sheets = openpyxl.load_workbook(excel_file, read_only=True).worksheets
    visible_sheets = []
    for sheet in sheets:
        if sheet.sheet_state != 'hidden':
            visible_sheets.append(sheet.title)
    if len(visible_sheets) != 1:
        raise Exception('The Excel file should contain only one visible sheet')
    return visible_sheets[0]


def get_excel_file():
    folder_path = Path('/', *os.getenv('REPORT_FOLDER').split(',')).resolve()
    files = os.listdir(folder_path)
    target_file = [Path(folder_path, file) for file in files if file.__contains__('User List')]
    if not target_file:
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), 'related user list excel file')
    else:
        return target_file[0]


def clean_email(dataframe, email_column):
    dataframe[email_column] = dataframe[email_column].str.lower()
    dataframe[email_column] = dataframe[email_column].str.strip()
    return dataframe


def send_notification():
    load_dotenv()
    user_column = os.getenv('MFA_REPORT_USER_COLUMN')
    line_manager_column = os.getenv('MFA_REPORT_LINE_MANAGER_COLUMN')
    excel_file = get_excel_file()
    visible_sheet_name = get_visible_sheet_name(excel_file)
    df = pd.read_excel(excel_file, sheet_name=visible_sheet_name)

    df = clean_email(df, user_column)
    df = clean_email(df, line_manager_column)
    df = df.replace({np.nan: None, pd.NaT: None, '': None})
    validate_result = validate_info(df, user_column)
    if validate_result:
        issue_data = generate_issue_data(df, validate_result, user_column)
        Emails('mfa_error').send_mfa_error_email(issue_data.to_html(index=False), attach_excel(excel_file))
    else:
        for index, row in df.iterrows():
            Emails('mfa_request').send_mfa_request_email(row['User'], row['Line Manager'])

    sent_summary_info = generate_summary_data()
    Emails('mfa_summary').send_mfa_summary_email(sent_summary_info, attach_excel(excel_file))


if __name__ == '__main__':
    send_notification()
