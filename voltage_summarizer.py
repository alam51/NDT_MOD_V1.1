import datetime
import os

import pandas as pd
from dateutil.parser import parse as date_parse


def process_dod(df: pd.DataFrame, _date_str) -> pd.DataFrame:
    df.columns = ['SS', 'Voltage_Max', 'Max_Datetime', 'Voltage_Min', 'Min_Datetime']
    max_datetime_list = [_date_str + str(hr) for hr in df.loc[:, 'Max_Datetime']]
    df.loc[:, 'Max_Datetime'] = pd.to_datetime(max_datetime_list,
                                               dayfirst=False, yearfirst=True,
                                               infer_datetime_format=True,
                                               errors='coerce')
    min_datetime_list = [_date_str + str(hr) for hr in df.loc[:, 'Min_Datetime']]
    df.loc[:, 'Min_Datetime'] = pd.to_datetime(min_datetime_list, dayfirst=False, yearfirst=True,
                                               infer_datetime_format=True,
                                               errors='coerce')

    return df.dropna()


if __name__ == '__main__':
    t1 = datetime.datetime.now()
    main_df = pd.DataFrame()
    file_path = r'G:\My Drive\Monthly_Report\2022\January\DOD1'
    files = os.listdir(file_path)
    for file in files:
        if (file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm')) and '~' not in file:
            try:
                date = date_parse(file, dayfirst=True, fuzzy=True)
                date_str = str(date)[:11]
                dod_file_name = os.path.join(file_path, file)

                dod_df1 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='B:F', header=[0],
                                        index_col=None).dropna()

                dod_df2 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='G:K', header=[0],
                                        index_col=None).dropna()

                df1 = process_dod(dod_df1, date_str)
                df2 = process_dod(dod_df2, date_str)
                main_df = pd.concat([main_df, df1, df2], ignore_index=True)
                print(f'Finished file: {file}')
            except:
                print(f'---Failed in file: {file}---')

    output_path = os.path.join(file_path, 'voltage_summary.xlsx')
    main_df['Max_Datetime'] = main_df['Max_Datetime'].apply(lambda x: x.replace(tzinfo=None))
    main_df['Min_Datetime'] = main_df['Min_Datetime'].apply(lambda x: x.replace(tzinfo=None))
    main_df.to_excel(output_path)
    print(f'Voltage Summary Saved in {output_path}')
    print(f'Time taken {datetime.datetime.now() - t1}')
