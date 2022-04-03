import pandas as pd

dod_file_name = r'dod.xlsm'
# col_range1 =
dod_df1 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='B:F', header=[0], index_col=None).dropna()
dod_df2 = pd.read_excel(dod_file_name, sheet_name='Volt', skiprows=4, usecols='G:K', header=[0], index_col=None).dropna()
dod_df2.columns = dod_df1.columns

def add_date(df: pd.DataFrame)->pd.DataFrame:


a = 5
