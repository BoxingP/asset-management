import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from emails.emails import Emails


def filter_nonempty_data(dataframe, column):
    return dataframe[dataframe[column].notna() & (dataframe[column] != '')].reset_index(drop=True)


def send_notification():
    load_dotenv()
    output_folder_path = Path('/', *os.getenv('OUTPUT_FOLDER').split(',')).resolve()
    excel_file = pd.ExcelFile(Path(output_folder_path, os.getenv('ASSET_REPORT_OUTPUT')))
    df = pd.read_excel(excel_file, sheet_name=os.getenv('ASSET_REPORT_DATA_SHEET'))
    key_column = os.getenv('ASSET_REPORT_PRIMARY_KEY')
    model_column = os.getenv('ASSET_REPORT_MODEL_COLUMN')
    notification_column = os.getenv('ASSET_REPORT_SEND_NOTIFICATION_TO_COLUMN')

    df.replace({None: '', pd.NA: '', float('nan'): ''}, inplace=True)
    df = filter_nonempty_data(df, notification_column)

    grouped_df = df.groupby(notification_column).apply(lambda group: group.loc[:, [key_column, model_column]])
    grouped_df.rename(columns={key_column: 'IT设备编号', model_column: '型号'}, inplace=True)
    grouped_df['是否在用此设备（是/否）'] = ''
    for email, info in grouped_df.groupby(level=0):
        info.reset_index(drop=True, inplace=True)
        info_html = info.to_html(index=False)
        Emails('asset').send_asset_email(email, info_html)


if __name__ == '__main__':
    send_notification()
