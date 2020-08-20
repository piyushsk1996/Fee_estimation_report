from pathlib import Path
import pandas as pd
from StyleFrame import StyleFrame
from datetime import date
from datetime import datetime
from itertools import tee, islice, chain

# Setting Options for printing pandas dataframe
desired_width = 520
# Setting up Display length of dataframe
pd.set_option('display.width', desired_width)
# Setting up maximum no. of dataframe columns to display
pd.set_option('display.max_columns', 25)


def previous_and_next(some_iterable):
    prevs, items, nexts = tee(some_iterable, 3)
    prevs = chain([None], prevs)
    nexts = chain(islice(nexts, 1, None), [None])
    return zip(prevs, items, nexts)


# Getting Today's Date
today_date_obj = date.today()
today_date = today_date_obj.strftime('%Y-%m-%d')


def generate_excel(filename, foldername):
    # Setting up paths
    mypath = Path().absolute()
    excelfilename = mypath / foldername

    # Getting Excel writer
    excel_writer = StyleFrame.ExcelWriter(excelfilename / filename)

    # Writing Sheetname
    sheetname = "Fee Estimate Report"

    # Defining writers and getting worksheet object to write data to
    writer = pd.ExcelWriter(excel_writer, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheetname)
    writer.sheets[sheetname] = worksheet

    # Defining Format Structures
    date_format = workbook.add_format({
        'num_format': 'm/d/yyyy',
        'font_size': 11
    })

    # For Merging cells
    merge_format = workbook.add_format({
        'align': 'center',
        'border': 1,
        'valign': 'vcenter',
        'italic': True,
        'bold': True
    })

    # For header format
    header_format = workbook.add_format({
        'font_size': 11,
        'font': 'Calibri',
        'align': 'center',
        'bold': True,
        'valign': 'vcenter'
    })

    # For cell format
    cell_format = workbook.add_format({
        'font_size': 11,
        'font': 'Calibri',
        'align': 'left'
    })

    # Custom format for location name, primary insurance and secondary insurance
    custom_cell_format = workbook.add_format({
        'font_size': 11,
        'font': 'Calibri',
        'align': 'center',
        'bg_color': '#FFC000',
    })

    # Setting Column widths
    worksheet.set_column('A:A', 26.56)
    worksheet.set_column('B:B', 15.22, date_format)
    worksheet.set_column('C:C', 19.56)
    worksheet.set_column('D:D', 29.22)
    worksheet.set_column('E:E', 28.33)
    worksheet.set_column('F:F', 26.67)
    worksheet.set_column('G:G', 15.78)

    # Setting Row height
    worksheet.set_row(1, 167)

    # Writing user notice
    worksheet.merge_range('A1:E1', 'Note: user can collapse the detail columns if only total row is needed. When '
                                   'discussing amount due with patient, see column P or U as directed', merge_format)

    # Writing headers to the excel
    worksheet.write("B2", "Date of Appt", header_format)
    worksheet.write("C2", "Office Location", header_format)
    worksheet.write("D2", "Primary Insurance ", header_format)
    worksheet.write("E2", "Secondary Insurance", header_format)
    worksheet.write("F2", "Program", header_format)
    worksheet.write("G2", "Patient treatment", header_format)

    # Reading report
    df_data_auth_report = pd.read_excel("./automation_files/OMG 0702 Chemo Auth Report.xlsx")

    # Setting NAN values to zero for easier handling
    df_data_auth_report = df_data_auth_report.fillna(0)

    # Reading CSV with treatment information
    df_read_csv = pd.read_csv('./All_CSVs/' + today_date + '.csv')

    # Upper casing the patient name for merging operation
    df_data_auth_report['Patient Name'] = df_data_auth_report['Patient Name'].str.upper()
    # Merging Operation
    df_merged = df_data_auth_report.merge(df_read_csv, on='Patient Name', how='left')

    # Create list of rows to be deleted according to the given conditions
    # 1. Sodium Chloride
    # 2. Oral Drugs ( to be added)
    # 3. Famotidine (to be added. Scenario to be tested)
    # 4. Tylenol (to be added. Scenario to be tested)

    delete_rows = list()

    # iterating through rows to detect keywords
    for index, row in df_merged.iterrows():
        for col in df_merged.columns:
            if col == "Treatment":
                # Condition for deletion
                if "sodium chloride" in str(row[col]):
                    # Retrieving Index
                    store_index = index
                    # Appending index to the list
                    delete_rows.append(store_index)

    # Deleting rows one by according to the given index
    for i in delete_rows:
        df_merged = df_merged.drop(df_merged.loc[i, :])

    # Outputting the merged dataframe to csv to check whether key words have been removed
    df_merged.to_csv("1.csv")
    df_merged = df_merged.fillna(0)
    # Setting initial row number
    initial_row_no = 3

    # Looping through column Patient Name and writing each name to the excel sheet
    for name_tuple, date_value, primary_insurance, secondary_insurance, location_name, treatment_name in zip(
            previous_and_next(df_merged["Patient Name"]),
            df_merged["Time"],
            df_merged["Primary "
                      "Insurance "
                      "Name"],
            df_merged["Secondary "
                      "Insurance "
                      "Name"],
            df_merged["Location"],
            df_merged["Treatment"]):
        previous_name, name, nxt_name = name_tuple
        if previous_name != name:

            # Splitting Name in to first name and last name
            first_name = name.split(',')[1]
            last_name = name.split(',')[0]

            # Splitting report date
            report_date = date_value.split(' ')[0]
            date_time_obj = datetime.strptime(report_date, '%m-%d-%Y')
            report_date_val = date_time_obj.strftime('%d-%m-%Y')

            # Writing Names
            worksheet.write(initial_row_no, 0, first_name + ' ' + last_name)
            new_row_no = initial_row_no + 1

            worksheet.write(new_row_no, 0, first_name + ' ' + last_name)
            blank_row_no = new_row_no + 1

            # Adding a blank row
            worksheet.write_blank(blank_row_no, 0, None, cell_format)

            # Writing dates to the corresponding columns
            worksheet.write(initial_row_no, 1, report_date_val)
            worksheet.write(new_row_no, 1, report_date_val)
            worksheet.write_blank(blank_row_no, 1, None, cell_format)

            # Writing Location Name
            worksheet.write(new_row_no, 2, location_name, custom_cell_format)

            # Writing Primary Insurance Name
            if primary_insurance == 0:
                worksheet.write_blank(new_row_no, 3, None, custom_cell_format)
            else:
                worksheet.write(new_row_no, 3, primary_insurance, custom_cell_format)

            # Writing Secondary Insurance Name
            if secondary_insurance == 0:
                worksheet.write_blank(new_row_no, 4, None, custom_cell_format)
            else:
                worksheet.write(new_row_no, 4, secondary_insurance, custom_cell_format)

            worksheet.write_blank(new_row_no, 5, None, custom_cell_format)

            if treatment_name == 0:
                worksheet.write_blank(new_row_no, 6, None)
            else:
                print(name)
                print(treatment_name)
                treatment_name_part1 = treatment_name.split(' ')[0]
                try:
                    treatment_name_part2 = treatment_name.split(' ')[1]
                except Exception as e:
                    treatment_name_part2 = None

                df_codes = pd.read_excel("./automation_files/ASP Pricing Copy.xlsx")
                df_codes['Short Description'] = df_codes['Short Description'].str.upper()
                for index, row in df_codes.iterrows():
                    try:
                        if treatment_name_part1.upper() in row['Short Description']:
                            hcpcs_code = df_codes.loc[index, "HCPCS Code"]
                            print(hcpcs_code)

                        elif treatment_name_part2 is not None and treatment_name_part2 in row['Short Description']:
                            hcpcs_code = df_codes.loc[index, "HCPCS Code"]
                            print(hcpcs_code)
                    except Exception as e:
                        hcpcs_code = None

                worksheet.write(new_row_no - 1, 6, hcpcs_code, custom_cell_format)
        else:
            treatment_name_part1 = treatment_name.split(' ')[0]
            try:
                treatment_name_part2 = treatment_name.split(' ')[1]
            except Exception as e:
                treatment_name_part2 = None

            df_codes = pd.read_excel("./automation_files/ASP Pricing Copy.xlsx")
            for index, row in df_codes.iterrows():
                if treatment_name_part1 in row['Short Description']:
                    hcpcs_code = df_codes.loc[index, "HCPCS Code"]
                    print(hcpcs_code)

                elif treatment_name_part2 is not None and treatment_name_part2 in row['Short Description']:
                    hcpcs_code = df_codes.loc[index, "HCPCS Code"]
                    print(hcpcs_code)
            worksheet.write(initial_row_no - 4, 6, hcpcs_code, custom_cell_format)

        # Incrementing Counter
        initial_row_no += 3

    writer.save()


if __name__ == '__main__':
    # Calling function
    generate_excel()