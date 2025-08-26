import pandas as pd
import numpy as np
import re

def clean_dataframes(dataframes):
    """
    Cleans a list of company dataframes by removing unnecessary text,
    converting data types, and handling missing values.
    """
    for df in dataframes:
        # 1. Clean column names
        patterns_to_remove = r'\s*\(\s*IFRS\s*연결\s*\)|연간컨센서스보기'
        df.columns = [re.sub(patterns_to_remove, '', str(col)).strip() for col in df.columns]

        # 2. Clean the first column's text
        first_col_name = df.columns[0]
        if df[first_col_name].dtype == 'object':
            df[first_col_name] = df[first_col_name].str.replace('\xa0', ' ', regex=False).str.strip()
            df[first_col_name] = df[first_col_name].str.replace('펼치기', '', regex=False).str.strip()
            df[first_col_name] = df[first_col_name].str.replace('(수익)', '', regex=False)
            df[first_col_name] = df[first_col_name].str.replace(r'\s*\(\s*IFRS\s*연결\s*\)', '', regex=True).str.strip()
            df.loc[df[first_col_name] == '', first_col_name] = np.nan

        # 3. Convert data columns to numeric
        for col in df.columns[1:]:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # 4. Drop unnecessary rows
        df.dropna(subset=[first_col_name], inplace=True)
        df.dropna(subset=df.columns[1:], how='all', inplace=True)
    return dataframes