import datetime

import pandas as pd
from dateutil.parser import parse as date_parse

dod_file_name = r'dod.xlsm'
# col_range1 =
dod_df1 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='B:F', header=[0],
                        index_col=None).dropna()

dod_df2 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='G:K', header=[0],
                        index_col=None).dropna()
dod_df2.columns = dod_df1.columns
date = date_parse('21 Feb 2022', dayfirst=True, fuzzy=True)
date_str = str(date)[:11]


def process_dod(df: pd.DataFrame, _date_str) -> pd.DataFrame:
    df.columns = ['SS', 'Voltage_Max', 'Max_Datetime', 'Voltage_Min', 'Min_Datetime']
    max_datetime_list = [_date_str + str(hr) for hr in df.loc[:, 'Max_Datetime']]
    df.loc[:, 'Max_Datetime'] = pd.to_datetime(max_datetime_list, dayfirst=False, yearfirst=True,
                                               infer_datetime_format=True, errors='coerce')
    min_datetime_list = [_date_str + str(hr) for hr in df.loc[:, 'Min_Datetime']]
    df.loc[:, 'Min_Datetime'] = pd.to_datetime(min_datetime_list, dayfirst=False, yearfirst=True,
                                               infer_datetime_format=True, errors='coerce')

    return df.dropna()


df1 = process_dod(dod_df1, date_str)
a = 5
