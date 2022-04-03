import pandas as pd


class VoltageDF:
    def __init__(self, df: pd.DataFrame):
        self.df1 = pd.read_excel(df, sheet_name='Volt', skiprows=4, usecols='B:F', header=[0], index_col=None).dropna()
        self.df2 = pd.read_excel(df, sheet_name='Volt', skiprows=4, usecols='G:K', header=[0], index_col=None).dropna()
        self.df2.columns = self.df1.columns

    def add_date(self):
