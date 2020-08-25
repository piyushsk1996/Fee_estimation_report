from pathlib import Path
import pandas as pd
from StyleFrame import StyleFrame
from datetime import date
from datetime import datetime
from itertools import tee, islice, chain
import numpy as np

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

    formula_cell_format = workbook.add_format({
        'font_size': 11,
        'font': 'Calibri',
        'align': 'center',
        'bg_color': '#B2B2B2',
    })

    # Setting Column widths
    worksheet.set_column('A:A', 26.56)
    worksheet.set_column('B:B', 15.22, date_format)
    worksheet.set_column('C:C', 19.56)
    worksheet.set_column('D:D', 29.22)
    worksheet.set_column('E:E', 28.33)
    worksheet.set_column('F:F', 26.67)
    worksheet.set_column('G:G', 15.78)
    worksheet.set_column('H:H', 17.11)
    worksheet.set_column('I:I', 15.56)

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
    worksheet.write("H2", "Treatment quantity", header_format)
    worksheet.write("I2", "Standard quantity", header_format)

    # Reading report
    df_data_auth_report = pd.read_excel("./All_Files/OMG 0826 Chemo Auth Report.xlsx")
    # Upper casing patient name column for matching
    df_data_auth_report['Patient Name'] = df_data_auth_report['Patient Name'].str.upper()
    # Setting NAN values to zero for easier handling
    df_data_auth_report = df_data_auth_report.fillna(0)

    # Reading J and Q code data from source
    df_J_Q = pd.read_excel("./All_Files/J and Q codes.xlsx")
    print(df_J_Q)
    # Writing J and Q codes data to current excel
    df_J_Q.to_excel(writer, index=False, sheet_name="J and Q Codes par amount", startrow=0, startcol=0)

    # Reading CSV with treatment information
    df_read_csv = pd.read_csv('./All_CSVs/' + today_date + '.csv')

    # Outputting the merged dataframe to csv to check whether key words have been removed

    df_read_csv.to_csv("1.csv")

    # Setting initial row number
    initial_row_no = 3
    # calculating number of rows required for each patient
    patient_dict = dict()
    df_read_csv = df_read_csv.fillna(0)
    for name, treatment_name, dosage in zip(df_read_csv["Patient Name"], df_read_csv["Treatment"],
                                            df_read_csv["Dosage"]):
        if treatment_name != 0:

            treatment = treatment_name.split(' ')[0]
            df_codes = pd.read_excel("./All_Files/ASP Pricing Copy.xlsx")
            df_codes['Short Description'] = df_codes['Short Description'].str.upper()

            for index, row in df_codes.iterrows():
                if treatment.upper() in row['Short Description']:
                    hcpcs_code = df_codes.loc[index, "HCPCS Code"]

                    temp_dict = {hcpcs_code: dosage}

                    if name != 0:
                        first_name = name.split(',')[1].strip().split(' ')[0]
                        last_name = name.split(',')[0].strip()
                        try:
                            middle_name = name.split(',')[1].strip().split(' ')[1]
                        except Exception as e:
                            middle_name = ''
                        if middle_name == '':
                            patient_dict.setdefault(last_name + ', ' + first_name, []).append(temp_dict)
                            # patient_dict[last_name + ', ' + first_name][hcpcs_code] = dosage
                        else:
                            patient_dict.setdefault(last_name + ', ' + first_name + ', ' + middle_name, []).append(
                                temp_dict)

    print(patient_dict)
    # Looping through column Patient Name and writing each name to the excel sheet
    for name, date_value, primary_insurance, secondary_insurance, location_name in zip(
            df_data_auth_report["Patient Name"],
            df_data_auth_report["Time"],
            df_data_auth_report["Primary "
                                "Insurance "
                                "Name"],
            df_data_auth_report["Secondary "
                                "Insurance "
                                "Name"],
            df_data_auth_report["Location"], ):

        print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
        # Splitting Name in to first name and last name
        first_name = name.split(',')[1].strip()
        last_name = name.split(',')[0]

        df_program = pd.read_excel("./All_Files/Program Pts (2).xlsx")
        df_program = df_program.replace(np.nan, '', regex=True)

        df_program_dict = dict()
        df_program_dict['program_name'] = None

        try:

            for index_val, row_val in df_program.iterrows():

                if last_name == str(row_val['Last Name']).strip().upper() and first_name == str(
                        row_val['First Name'].strip().upper()):
                    program_index = index_val
                    program_name = df_program.loc[program_index, 'Program']
                    df_program_dict['program_name'] = program_name
                    print(df_program_dict)
        except Exception as e:
            print("Program not found")
            df_program_dict['program_name'] = None

        if name in patient_dict.keys():
            span_patient = len(patient_dict[name])

        initial_row_no = initial_row_no + span_patient

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

        worksheet.write(new_row_no, 5, df_program_dict['program_name'], custom_cell_format)

        cpt_code_row_no = new_row_no - 1
        # writing cpt codes
        for key, values in patient_dict.items():
            key_first_name = key.split(',')[1].strip().split(' ')[0]
            key_last_name = key.split(',')[0].strip()

            if key_first_name == first_name and key_last_name == last_name:
                for cpt_code in values:
                    for cpt_key, cpt_val in cpt_code.items():
                        worksheet.write(cpt_code_row_no, 6, cpt_key, custom_cell_format)
                        worksheet.write(cpt_code_row_no, 7, cpt_val, custom_cell_format)
                        worksheet.write_formula(cpt_code_row_no, 8,
                                                "=VLOOKUP(G" + str(cpt_code_row_no) + ",'J and Q Codes par amount'!A:C,"
                                                                                      "2,FALSE)",
                                                formula_cell_format)
                    cpt_code_row_no -= 1

        # Incrementing Counter
        initial_row_no += 3

    writer.save()


if __name__ == '__main__':
    # Calling function
    generate_excel(filename='MHP FEE ESTIMATE Excel Report.xlsx', foldername='Excel_files')
