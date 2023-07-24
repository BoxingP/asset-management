import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from emails.emails import Emails


def get_report_path():
    report_folder_path = Path('/', *os.getenv('OUTPUT_FOLDER').split(',')).resolve().absolute()
    files = report_folder_path.glob('*.xlsx')
    return [file.absolute() for file in files]


def send_notification():
    load_dotenv()
    report_path = get_report_path()
    excel_file = pd.ExcelFile(report_path[0])
    df = pd.read_excel(excel_file, sheet_name=os.getenv('REPORT_SHEET'))

    df.replace({None: '', pd.NA: '', float('nan'): ''}, inplace=True)
    df['Most recent discovery'] = df['Most recent discovery'].astype(str).replace('NaT', '')

    grouped_df = df.groupby('Notification Email').apply(lambda group: group.loc[:, ['SN号', 'Manufacturer']])
    grouped_df.rename(columns={'SN号': 'IT设备编号', 'Manufacturer': '型号'}, inplace=True)
    grouped_df['是否在用此设备（是/否）'] = ''
    for email, info in grouped_df.groupby(level=0):
        info.reset_index(drop=True, inplace=True)
        info_html = info.to_html(index=False)
        Emails().send_email(email, info_html)


if __name__ == '__main__':
    send_notification()
