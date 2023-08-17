import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from emails.emails import Emails
from utils.email_logger import EmailSendingLogger


def validate_id_uniqueness(index, dataframe, column):
    is_uniqueness = dataframe[column].nunique() == 1
    if not is_uniqueness:
        return {'index': index, 'reason': '员工号不唯一'}


def validate_time_uniqueness(index, dataframe, column):
    is_uniqueness = dataframe[column].dt.date.nunique() == 1
    if not is_uniqueness:
        return {'index': index, 'reason': '交还日期不唯一'}


def validate_sn_uniqueness(index, dataframe, column):
    is_uniqueness = len(dataframe[column]) == len(dataframe[column].unique())
    if not is_uniqueness:
        return {'index': index, 'reason': '设备序列号有重复'}


def validate_info(dataframe, id_column, date_column, sn_column):
    validate_result = []
    for index, info in dataframe.groupby(level=0):
        validate_result.append(validate_id_uniqueness(index, info, id_column))
        validate_result.append(validate_time_uniqueness(index, info, date_column))
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


def generate_issue_data(dataframe, validate_result, key_column, columns_to_remove):
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
    return result_df.drop(columns=columns_to_remove)


def generate_summary_data():
    df = pd.DataFrame(EmailSendingLogger().get_log_data())
    df[['name', 'emp_id']] = df['subject'].str.extract(r'IT资产归还提醒\((\D+)(\d+)\)')
    df['success'] = df['success'].replace({True: 'Y', False: 'N'})
    df = df[['time', 'emp_id', 'name', 'recipient', 'subject', 'success', 'error_message']]
    df.rename(columns={'time': '发送时间', 'emp_id': '员工号', 'name': '员工中文名', 'recipient': '收件邮箱',
                       'subject': '邮件标题', 'success': '是否发送成功', 'error_message': '详细信息'}, inplace=True)
    return df


def send_notification():
    load_dotenv()
    output_folder_path = Path('/', *os.getenv('OUTPUT_FOLDER').split(',')).resolve()
    excel_file = pd.ExcelFile(Path(output_folder_path, os.getenv('RETURN_REPORT_OUTPUT')))
    key_column = os.getenv('RETURN_REPORT_PRIMARY_KEY')
    id_column = os.getenv('RETURN_REPORT_ID_COLUMN')
    name_column = os.getenv('RETURN_REPORT_NAME_COLUMN')
    email_column = os.getenv('RETURN_REPORT_EMAIL_COLUMN')
    model_column = os.getenv('RETURN_REPORT_MODEL_COLUMN')
    sn_column = os.getenv('RETURN_REPORT_SN_COLUMN')
    state_column = os.getenv('RETURN_REPORT_STATE_COLUMN')
    date_column = os.getenv('RETURN_REPORT_DATE_COLUMN')
    df = pd.read_excel(excel_file, sheet_name=os.getenv('RETURN_REPORT_OUTPUT_SHEET'), dtype={id_column: str})

    copy_df = df
    copy_df[key_column] = copy_df[key_column].str.lower()
    grouped_df = copy_df.groupby(key_column). \
        apply(lambda group: group.loc[:, [id_column, name_column, email_column, model_column, sn_column, state_column,
                                          date_column]])

    validate_result = validate_info(grouped_df, id_column, date_column, sn_column)
    if validate_result:
        issue_data = generate_issue_data(df, validate_result, key_column, [state_column, date_column])
        Emails('return_error').send_return_error_email(issue_data.to_html(index=False))
    else:
        grouped_df.rename(columns={model_column: '设备型号', sn_column: '设备序列号', state_column: '归还状态'}, inplace=True)
        for index, info in grouped_df.groupby(level=0):
            Emails('return').send_return_email(info[id_column][0], info[name_column][0], info[email_column][0],
                                               info[date_column][0], info.to_html(index=False))

        sent_summary_df = generate_summary_data()
        Emails('return_summary').send_return_summary_email(sent_summary_df.to_html(index=False))


if __name__ == '__main__':
    send_notification()
