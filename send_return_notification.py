import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from emails.emails import Emails


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

    df[key_column] = df[key_column].str.lower()
    grouped_df = df.groupby(key_column). \
        apply(lambda group: group.loc[:, [id_column, name_column, email_column, model_column, sn_column, state_column,
                                          date_column]])
    grouped_df.rename(columns={model_column: '设备型号', sn_column: '设备序列号', state_column: '归还状态'}, inplace=True)
    for index, info in grouped_df.groupby(level=0):
        info[date_column] = pd.to_datetime(info[date_column])
        info_html = info[['设备型号', '设备序列号', '归还状态']].drop_duplicates().to_html(index=False)
        Emails('return').send_return_email(index, info[name_column][0], info[email_column][0], info[date_column].min(),
                                           info_html)


if __name__ == '__main__':
    send_notification()
