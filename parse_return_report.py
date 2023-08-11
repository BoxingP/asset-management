import datetime
import os
from pathlib import Path

import pandas as pd
import wcwidth
from dotenv import load_dotenv


def export_dataframe_to_excel(writer, dataframe, sheet_name):
    workbook = writer.book
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name].clear()
    else:
        sheet = workbook.add_worksheet(sheet_name)
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#5B9BD5',
        'font_color': '#FFFFFF'
    })
    fmt_time = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    row = 1
    for i, row_data in dataframe.iterrows():
        for col_idx, col_value in enumerate(row_data):
            if pd.isna(col_value):
                sheet.write(row, col_idx, None)
            elif isinstance(col_value, pd.Timestamp):
                sheet.write_datetime(row, col_idx, col_value.to_pydatetime(), fmt_time)
            else:
                sheet.write(row, col_idx, col_value)
        row += 1
    worksheet = writer.sheets[sheet_name]
    columns_width = [max(len(str(col)), wcwidth.wcswidth(col)) + 4 for col in dataframe.columns]
    for col_idx, col_name in enumerate(dataframe.columns):
        worksheet.set_column(col_idx, col_idx, columns_width[col_idx])
        worksheet.write(0, col_idx, col_name, header_format)


def filter_resignation_date(dataframe, target_month, target_year=datetime.datetime.now().year):
    dataframe['离职日期'] = pd.to_datetime(dataframe['离职日期'])
    return dataframe[(dataframe['离职日期'].dt.year == target_year) & (dataframe['离职日期'].dt.month == target_month)]


def filter_return_state(dataframe, target_state):
    return dataframe[dataframe['是否归还或转移日期'].str.contains(target_state)]


def get_file_path_by_filename(paths, target_filename):
    for path in paths:
        filename = path.name
        if filename == target_filename:
            return path
    return None


def parse_report():
    load_dotenv()
    report_folder_path = Path('/', *os.getenv('REPORT_FOLDER').split(',')).resolve()
    excel_file = pd.ExcelFile(Path(report_folder_path, os.getenv('RETURN_REPORT_NAME')))
    origin_df = pd.read_excel(excel_file, sheet_name=os.getenv('RETURN_REPORT_DATA_SHEET'))

    selected_columns = os.getenv('RETURN_REPORT_COLUMN').split(',')
    df = origin_df.loc[:, selected_columns]
    df = filter_return_state(df, os.getenv('RETURN_REPORT_STATE'))
    df = filter_resignation_date(df, target_month=int(os.getenv('RETURN_REPORT_MONTH')))

    output_folder_path = Path('/', *os.getenv('OUTPUT_FOLDER').split(',')).resolve()
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
    with pd.ExcelWriter(Path(output_folder_path, os.getenv('RETURN_REPORT_OUTPUT')), engine='xlsxwriter') as writer:
        export_dataframe_to_excel(writer, df, 'test')


if __name__ == '__main__':
    parse_report()
