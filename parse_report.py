import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv


def export_dataframe_to_excel(writer, dataframe, sheet_name, columns_width):
    workbook = writer.book
    if sheet_name in workbook.sheetnames:
        workbook[sheet_name].clear()
    else:
        workbook.add_worksheet(sheet_name)
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#5B9BD5',
        'font_color': '#FFFFFF'
    })
    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    for i, width in enumerate(columns_width):
        worksheet.set_column(i, i, width)
    for col_idx, col_name in enumerate(dataframe.columns):
        if col_idx < len(columns_width):
            worksheet.write(0, col_idx, col_name, header_format)


def group_data(dataframe, group_column, ignore_column, notification_column):
    filtered_df = dataframe.loc[dataframe[f'{group_column}邮箱'].notna() & (dataframe[f'{group_column}邮箱'] != '')]
    grouped_df = filtered_df.groupby(f'{group_column}邮箱', as_index=False).apply(
        lambda group: group.dropna(subset=['SN号'])).reset_index(drop=True)
    grouped_df.rename(columns={f'{group_column}邮箱': notification_column, f'{group_column} Band': 'Band'}, inplace=True)
    grouped_df['Got from'] = group_column
    grouped_df.drop(columns=[f'{ignore_column}邮箱', f'{ignore_column} Band'], inplace=True)
    return grouped_df


def get_report_path():
    report_folder_path = Path('/', *os.getenv('REPORT_FOLDER').split(',')).resolve().absolute()
    files = report_folder_path.glob('*.xlsx')
    return [file.absolute() for file in files]


def parse_report():
    load_dotenv()
    report_path = get_report_path()
    excel_file = pd.ExcelFile(report_path[0])
    origin_df = pd.read_excel(excel_file, sheet_name=os.getenv('REPORT_DATA_SHEET'))
    key_column = os.getenv('REPORT_PRIMARY_KEY')
    notification_column = os.getenv('REPORT_SEND_NOTIFICATION_TO_COLUMN')

    selected_columns = os.getenv('REPORT_COLUMN').split(',')
    df = origin_df.loc[:, selected_columns]
    df = df[df[key_column].notna() & (df[key_column] != '')].reset_index(drop=True)
    owner_group = group_data(df, 'Owner', 'User', notification_column)
    owner_group_sn = owner_group[key_column].reset_index(drop=True)
    new_df = df[~df[key_column].isin(owner_group_sn)].reset_index(drop=True)
    user_group = group_data(new_df, 'User', 'Owner', notification_column)
    processed_df = pd.concat([owner_group, user_group], ignore_index=True)
    processed_df = processed_df[
        (processed_df['Band'] < int(os.getenv('REPORT_IGNORED_BAND'))) | processed_df['Band'].isna()].reset_index(
        drop=True)

    selected_columns = [key_column, notification_column]
    final_df = origin_df.merge(processed_df[selected_columns], on=key_column, how='left')
    final_df[notification_column] = final_df[notification_column].fillna('')

    with pd.ExcelWriter(report_path[0], engine='xlsxwriter') as writer:
        export_dataframe_to_excel(writer, final_df, os.getenv('REPORT_RESULT_SHEET'),
                                  [11, 20, 16, 16, 30, 25, 30, 13, 13, 30])


if __name__ == '__main__':
    parse_report()
