import random
import string
import pandas as pd
import xlsxwriter
from datetime import date

today_date = date.today()


def vlookup_in_python(df, df1):
    # Merge category data by URL
    df2 = df.merge(df1, on='URL', how='left')
    df2.to_excel('new.xlsx', sheet_name="3")
    return True


# Check if a value is present within another column
def present_in(df, df1):
    # This function gets applied to each value within
    def is_in_column(data, values):
        if data in values:
            return data
        else:
            return False

    df['Present_in_other'] = df['URL'].apply(is_in_column, values=df1['URL'].tolist())
    return df




df, df1 = pd.read_excel("./automation_files/April 2020 ASP Pricing File 03022020.xls", sheet_name="1"), pd.read_excel("./All_CSVs/" + today_date +".csv", sheet_name="2")