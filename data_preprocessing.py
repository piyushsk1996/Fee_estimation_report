import camelot
import pandas as pd
import os
import numpy as np
from datetime import date

# Setting Options for printing pandas dataframe
desired_width = 520
# Setting up Display length of dataframe
pd.set_option('display.width', desired_width)
# Setting up maximum no. of dataframe columns to display
pd.set_option('display.max_columns', 25)


# Reading pdf data and pre-processing it with given conditions
def pre_process_data():
    """
    :return: status of csv creation
    """
    # Creating Empty Dataframe
    final_df = pd.DataFrame()

    # calculating todays date and setting it to y-m-d format
    today_date = date.today()
    today_date = today_date.strftime('%Y-%m-%d')

    # Calculate no. pdf parsed
    i = 0
    # Traverse through directory using os.walk
    for root, dirs, files in os.walk("./All_Pdf_Files/"):
        # Traversing through all pdf files
        for file in files:
            i += 1
            print(file)
            # Getting Tables using camelot and setting options line scale=40 and shift text = ['']
            tables = camelot.read_pdf("./All_Pdf_Files/" + file, line_scale=40, shift_text=[' '],
                                      layout_kwargs={'detect_vertical': False})
            # Maximum range of 4 to cover all tables
            for i in range(4):
                try:

                    # Converting tables to dataframe and indexing them using counter i
                    test_df = tables[i].df

                    if i == 0:
                        # Retriving patient name
                        patient_name = str(test_df.loc[0, 5]).replace("Patient:\n", "").split('\n')[0]

                    if "PRE-TREATMENT" == str(test_df.loc[0, 0]) or "TREATMENT" == str(test_df.loc[0, 0]):

                        # print(str(test_df.loc[0, 0].split()))
                        # Initialzing empty dictionary to store treatment data
                        treatment_dict = dict()
                        # Getting Patient Name:
                        treatment_dict.setdefault("Patient Name", []).append(patient_name)
                        # Getting Patient ID
                        treatment_dict.setdefault("Patient ID", []).append(file.split('.')[0])

                        # Getting Type of treatment
                        type_of_treatment = test_df.loc[0, 0]
                        # Appending type of treatment to dictionarty treament_dict
                        treatment_dict.setdefault("Type of treatment", []).append(type_of_treatment)

                        if "magnesium sulfate" in test_df[1].values:
                            test_df[0] = test_df[0] + test_df[1]
                            test_df[1] = test_df[2]

                        test_df.replace('', 0, inplace=True)

                        for index, row in test_df.iterrows():

                            if index != 0 and str(row[0]) != str(0):
                                if "sodium chloride" not in str(row[0]):
                                    if "magnesium" not in str(row[0]):
                                        if "potassium" not in str(row[0]):
                                            if "immune" in str(row[0]):
                                                treatment = str(row[0])
                                            else:
                                                treatment = str(row[0]).split(' ')[0]
                                            if ")" not in treatment:
                                                if "Delivery" not in treatment:
                                                    if "Dosed" not in treatment:
                                                        if "LAR" not in treatment:
                                                            if "(" not in treatment:
                                                                treatment_dict.setdefault("Treatment", []).append(
                                                                    treatment)

                                                    dosage = str(row[1]).split('\n')[0]
                                                    if "(d" not in dosage:
                                                        treatment_dict.setdefault("Dosage", []).append(dosage)

                        # Constructing dataframe from dictionary values
                        df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in treatment_dict.items()]))

                        # Concatenating dataframe to final_df and reassigning it to add multiple dataframes into one
                        # dataframe
                        final_df = pd.concat([final_df, df], axis=0)
                # Catch the exception
                except Exception as e:
                    print("Couldn't find more tables")
    # Filling NAN values with zeros
    final_df = final_df.fillna(0)
    for index, row in final_df.iterrows():
        if "mL/hr" in str(row["Dosage"]):
            row["Dosage"] = 0
            row["Treatment"] = 0
    # Reseting index to get new index
    final_df.reset_index(drop=True, inplace=True)
    # Dropping rows from dataframe based on given condition
    # Dropping rows where dosage and treatment is null
    final_df = final_df.drop(final_df[(final_df.Treatment == 0) & (final_df.Dosage == 0)].index)
    # Replacing zero values to nan
    final_df.replace(0, np.nan, inplace=True)
    # Exporting dataframe to csv
    final_df.to_csv('./All_CSVs/' + today_date + '.csv', index=False)
    print("No. of files parsed : ", len(files))
    return "Dataframe cleaned and Exported in CSV format"


if __name__ == '__main__':
    # Calling function
    pre_process_data()
