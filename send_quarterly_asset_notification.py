import errno
import os
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

from emails.emails import Emails
from utils.email_logger import EmailSendingLogger


def generate_summary_data(dataframe, name_column, email_column):
    df = pd.DataFrame(EmailSendingLogger().get_log_data())
    df = df.merge(dataframe, left_on='recipient', right_on=email_column, how='left')
    df['success'] = df['success'].replace({True: 'Y', False: 'N'})
    df = df[['time', name_column, 'recipient', 'subject', 'success', 'error_message']]
    df.rename(columns={'time': '发送时间', name_column: '员工中文名', 'recipient': '收件邮箱',
                       'subject': '邮件标题', 'success': '是否发送成功', 'error_message': '详细信息'}, inplace=True)
    soup = BeautifulSoup(df.to_html(index=False), 'html.parser')
    rows = soup.find_all('tbody')[0].find_all('tr')
    for row in rows:
        success_cell = row.find_all('td')[4]
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


def validate_name_uniqueness(index, dataframe, column):
    is_uniqueness = dataframe[column].nunique() == 1
    if not is_uniqueness:
        return {'index': index, 'reason': '员工名字不唯一'}


def validate_sn_uniqueness(index, dataframe, column):
    is_uniqueness = len(dataframe[column]) == len(dataframe[column].unique())
    if not is_uniqueness:
        return {'index': index, 'reason': '设备序列号有重复'}


def validate_info(dataframe, name_column, sn_column):
    validate_result = []
    for index, info in dataframe.groupby(level=0):
        validate_result.append(validate_name_uniqueness(index, info, name_column))
        validate_result.append(validate_sn_uniqueness(index, info, sn_column))
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


def clean_sn(dataframe, sn_column):
    dataframe[sn_column] = dataframe[sn_column].str.upper()
    dataframe[sn_column] = dataframe[sn_column].str.strip()
    return dataframe


def clean_email(dataframe, email_column):
    dataframe[email_column] = dataframe[email_column].str.lower()
    dataframe[email_column] = dataframe[email_column].str.strip()
    return dataframe


def get_excel_file():
    folder_path = Path('/', *os.getenv('REPORT_FOLDER').split(','),
                       *os.getenv('QUARTERLY_ASSET_REPORT_FOLDER').split(',')).resolve()
    files = os.listdir(folder_path)
    excel_files = [Path(folder_path, file) for file in files if file.endswith('.xlsx') or file.endswith('.xls')]
    if not excel_files:
        raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), 'related return excel file')
    else:
        return excel_files[0]


def generate_name_email_mapping(dataframe, name_column, email_column):
    df = dataframe.loc[:, [name_column, email_column]]
    df = df.drop_duplicates()
    return df


def send_notification():
    load_dotenv()
    email_column = os.getenv('QUARTERLY_ASSET_REPORT_EMAIL_COLUMN')
    name_column = os.getenv('QUARTERLY_ASSET_REPORT_NAME_COLUMN')
    model_column = os.getenv('QUARTERLY_ASSET_REPORT_MODEL_COLUMN')
    sn_column = os.getenv('QUARTERLY_ASSET_REPORT_SN_COLUMN')
    excel_file = get_excel_file()
    origin_df = pd.read_excel(excel_file, dtype={'资产号': str})
    selected_columns = os.getenv('QUARTERLY_ASSET_REPORT_COLUMN').split(',')
    df = origin_df.loc[:, selected_columns]
    df = df.dropna(how='all')

    df = clean_email(df, email_column)
    df = clean_sn(df, sn_column)
    grouped_df = df.groupby(email_column).apply(lambda group: group.loc[:, [name_column, model_column, sn_column]])

    validate_result = validate_info(grouped_df, name_column, sn_column)
    if validate_result:
        issue_data = generate_issue_data(df, validate_result, email_column)
        Emails('inventory_error').send_inventory_error_email(issue_data.to_html(index=False), attach_excel(excel_file))
    else:
        grouped_df.rename(columns={model_column: '型号', sn_column: '序列号/服务编号'}, inplace=True)
        for email, info in grouped_df.groupby(level=0):
            Emails('quarterly_asset').send_inventory_email(info[name_column][0], email,
                                                           info.loc[:, ['型号', '序列号/服务编号']].to_html(index=False))
        name_email_mapping = generate_name_email_mapping(df, name_column, email_column)
        sent_summary_info = generate_summary_data(name_email_mapping, name_column, email_column)
        Emails('inventory_summary').send_inventory_summary_email(sent_summary_info, attach_excel(excel_file))


if __name__ == '__main__':
    send_notification()
