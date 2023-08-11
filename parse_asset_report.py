import os
import re
from pathlib import Path

import pandas as pd
import wcwidth
from dotenv import load_dotenv


def export_dataframe_to_excel(writer, dataframe, sheet_name):
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
    columns_width = [max(len(str(col)), wcwidth.wcswidth(col)) + 4 for col in dataframe.columns]
    for col_idx, col_name in enumerate(dataframe.columns):
        worksheet.set_column(col_idx, col_idx, columns_width[col_idx])
        worksheet.write(0, col_idx, col_name, header_format)


def filter_manufacturer(dataframe, manufacturer_column, include_nan=False):
    manufacturers = '|'.join(os.getenv('ASSET_REPORT_MANUFACTURER').split(','))
    processed_df = dataframe[dataframe[manufacturer_column].str.contains(manufacturers, case=False, na=include_nan)]
    return processed_df


def filter_band(dataframe, notification_column, band_column, band: int):
    dataframe[band_column] = dataframe[band_column].apply(pd.to_numeric, errors='coerce')
    dataframe[band_column] = dataframe.groupby(notification_column)[band_column].transform('max')
    dataframe = dataframe[(dataframe[band_column] < band) | dataframe[band_column].isna()].reset_index(drop=True)
    return dataframe


def normalize_domain(email):
    return re.sub(r'@thermofisher\.com[^a-zA-Z]*', '@thermofisher.com', email)


def normalize_email(dataframe, notification_column):
    dataframe[notification_column] = dataframe[notification_column].str.strip().str.lower()
    dataframe[notification_column] = dataframe[notification_column].apply(normalize_domain)
    return dataframe


def group_data(dataframe, group_column, ignore_column, notification_column, band_column):
    filtered_df = filter_nonempty_data(dataframe, f'{group_column}邮箱')
    grouped_df = filtered_df.groupby(f'{group_column}邮箱', as_index=False).apply(
        lambda group: group.dropna(subset=['SN号'])).reset_index(drop=True)
    grouped_df.rename(columns={f'{group_column}邮箱': notification_column, f'{group_column} Band': band_column},
                      inplace=True)
    grouped_df['Got from'] = group_column
    grouped_df.drop(columns=[f'{ignore_column}邮箱', f'{ignore_column} Band'], inplace=True)
    return grouped_df


def filter_nonempty_data(dataframe, column):
    return dataframe[dataframe[column].notna() & (dataframe[column] != '')].reset_index(drop=True)


def parse_report():
    load_dotenv()
    report_folder_path = Path('/', *os.getenv('REPORT_FOLDER').split(',')).resolve()
    excel_file = pd.ExcelFile(Path(report_folder_path, os.getenv('ASSET_REPORT_NAME')))
    origin_df = pd.read_excel(excel_file, sheet_name=os.getenv('ASSET_REPORT_DATA_SHEET'))
    key_column = os.getenv('ASSET_REPORT_PRIMARY_KEY')
    notification_column = os.getenv('ASSET_REPORT_SEND_NOTIFICATION_TO_COLUMN')
    band_column = 'Band'
    statistical_column = os.getenv('ASSET_REPORT_STATISTICAL_COLUMN')

    selected_columns = os.getenv('ASSET_REPORT_COLUMN').split(',')
    df = origin_df.loc[:, selected_columns]
    df = filter_nonempty_data(df, key_column)
    owner_group_df = group_data(df, 'Owner', 'User', notification_column, band_column)
    owner_group_sn = owner_group_df[key_column].reset_index(drop=True)
    owner_group_remainder_df = df[~df[key_column].isin(owner_group_sn)].reset_index(drop=True)
    user_group_df = group_data(owner_group_remainder_df, 'User', 'Owner', notification_column, band_column)
    group_df = pd.concat([owner_group_df, user_group_df], ignore_index=True)

    group_df = normalize_email(group_df, notification_column)
    group_df = filter_band(group_df, notification_column, band_column, int(os.getenv('ASSET_REPORT_IGNORED_BAND')))

    selected_columns = [key_column, notification_column]
    final_df = origin_df.merge(group_df[selected_columns], on=key_column, how='left')
    final_df[notification_column] = final_df[notification_column].fillna('')
    final_df = filter_manufacturer(final_df, os.getenv('ASSET_REPORT_MANUFACTURER_COLUMN'))

    summary_df = final_df.groupby(notification_column)[key_column].size().reset_index()
    summary_df.rename(columns={key_column: statistical_column}, inplace=True)
    summary_df = summary_df.sort_values(by=statistical_column, ascending=False)
    total_row = pd.DataFrame(
        {notification_column: ['Total'], statistical_column: [summary_df[statistical_column].sum()]})
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)

    output_folder_path = Path('/', *os.getenv('OUTPUT_FOLDER').split(',')).resolve()
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
    with pd.ExcelWriter(Path(output_folder_path, os.getenv('ASSET_REPORT_OUTPUT')), engine='xlsxwriter') as writer:
        export_dataframe_to_excel(writer, final_df, os.getenv('ASSET_REPORT_DATA_SHEET'))
        export_dataframe_to_excel(writer, summary_df, os.getenv('ASSET_REPORT_SUMMARY_SHEET'))


if __name__ == '__main__':
    parse_report()
