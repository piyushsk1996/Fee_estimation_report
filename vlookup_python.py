import random
import string
import pandas as pd
import xlsxwriter


def create_dummy_data():
    words_1 = [[''.join(random.choices(string.ascii_uppercase, k=5)), "sheet1"] for i in range(0, 1000)]
    words_2 = [[''.join(random.choices(string.ascii_uppercase, k=5)), "sheet2"] for i in range(0, 1000)]

    # Append lists to URL column in df
    df = pd.DataFrame(words_1, columns=['URL', 'Item1'], index=None)
    df1 = pd.DataFrame(words_2, columns=['URL', 'Item2'], index=None)

    # Save to excel so you can benchmark on your machine
    writer = pd.ExcelWriter('vloop_test_data.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='1')
    df1.to_excel(writer, sheet_name='2')
    writer.save()


# Read an excel with two sheets into two dataframes
def load_data(name):
    df, df1 = pd.read_excel(name, sheet_name="1"), pd.read_excel(name, sheet_name="2")
    return df, df1


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


# Create data
create_dummy_data()


# # loading the data
# df1, df2 = load_data("vloop_test_data.xlsx")
#
# if vlookup_in_python(df1, df2):
#     df_result = present_in(df1, df2)
#
# print(df_result)

