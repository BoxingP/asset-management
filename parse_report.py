import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv


def create_folder(path):
    folder = Path('/', *path.split(',')).resolve().absolute()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def group_data(dataframe, group_column, ignore_column):
    filtered_df = dataframe.loc[dataframe[f'{group_column}邮箱'].notna() & (dataframe[f'{group_column}邮箱'] != '')]
    grouped_df = filtered_df.groupby(f'{group_column}邮箱', as_index=False).apply(
        lambda group: group.dropna(subset=['SN号'])).reset_index(drop=True)
    grouped_df.rename(columns={f'{group_column}邮箱': 'Notification Email', f'{group_column} Band': 'Band'}, inplace=True)
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
    origin_df = pd.read_excel(excel_file, sheet_name=os.getenv('REPORT_SHEET'))

    selected_columns = os.getenv('REPORT_COLUMN').split(',')
    df = origin_df.loc[:, selected_columns]
    df = df[df['SN号'].notna() & (df['SN号'] != '')].reset_index(drop=True)
    owner_group = group_data(df, 'Owner', 'User')
    owner_group_sn = owner_group['SN号'].reset_index(drop=True)
    new_df = df[~df['SN号'].isin(owner_group_sn)].reset_index(drop=True)
    user_group = group_data(new_df, 'User', 'Owner')
    processed_df = pd.concat([owner_group, user_group], ignore_index=True)
    processed_df = processed_df[(processed_df['Band'] <= 9) | processed_df['Band'].isna()].reset_index(drop=True)

    selected_columns = ['SN号', 'Notification Email']
    final_df = origin_df.merge(processed_df[selected_columns], on='SN号', how='left')
    final_df['Notification Email'] = final_df['Notification Email'].fillna('')

    output_folder = create_folder(os.getenv('OUTPUT_FOLDER'))
    final_df.to_excel(Path(output_folder, 'result.xlsx'), index=False)


if __name__ == '__main__':
    parse_report()
