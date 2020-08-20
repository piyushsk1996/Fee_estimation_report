import camelot
import pandas as pd
import os
import numpy as np

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

    # Traverse through directory using os.walk
    for root, dirs, files in os.walk("./All_Pdf_Files/"):
        # Traversing through all pdf files
        for file in files:
            print(file)
            # Getting Tables using camelot and setting options line scale=40 and shift text = ['']
            tables = camelot.read_pdf("./All_Pdf_Files/" + file, line_scale=40, shift_text=[''])
            # Maximum range of 4 to cover all tables
            for i in range(4):
                try:
                    # Converting tables to dataframe and indexing them using counter i
                    test_df = tables[i].df
                    if i == 0:
                        # Retriving patient name
                        patient_name = str(test_df.loc[0, 5]).replace("Patient:\n", "").split('\n')[0]

                    if "PRE-TREATMENT" == test_df.loc[0, 0] or "TREATMENT" == test_df.loc[0, 0]:
                        # Initialzing empty dictionary to store treatment data
                        treatment_dict = dict()
                        # Getting Patient Name:
                        treatment_dict.setdefault("Patient Name", []).append(patient_name)
                        # Getting Patient ID
                        # treatment_dict.setdefault("Patient ID", []).append(file.split('.')[0])

                        # Getting Type of treatment
                        type_of_treatment = test_df.loc[0, 0]

                        # Appending type of treatment to dictionarty treament_dict
                        treatment_dict.setdefault("Type of treatment", []).append(type_of_treatment)
                        # Looping through dataframe obtained from tables and processing it with conditions specified
                        for j in range(len(test_df)):
                            # Looping throgh dataframe columns to scan the dataframe
                            for k in range(len(test_df.columns)):
                                # Scanning column 0
                                if k == 0:
                                    # checking if any value exists
                                    if test_df.loc[j, k]:
                                        # checking if value is a type of treatment
                                        if test_df.loc[j, k] != treatment_dict["Type of treatment"][0]:
                                            # checking for new line charater and replacing it with a space
                                            if "\n" in test_df.loc[j, k]:
                                                treatment_string = test_df.loc[j, k].replace("\n", " ")
                                            else:
                                            # else assigning it as it is
                                                treatment_string = test_df.loc[j, k]
                                            # Checking for drug-condtions
                                            # 1. Sodium Chloride
                                            # 2. Oral drugs ( to be added)
                                            # 3. Famotidine ( to be added)
                                            # 4. Tylenol ( to be added )
                                            # Breaking if keyword is encountered while scanning dataframe
                                            if "sodium chloride" in treatment_string:
                                                break
                                            else:
                                                # Getting Treatment
                                                treatment_dict.setdefault("Treatment", []).append(treatment_string)
                                            # using trailing row variable to skip the rows
                                            trailin_row = j + 1
                                # Scanning column 1
                                elif k == 1:
                                    # checkig if value exists
                                    if test_df.loc[j, k]:
                                        # Using trailing row to skip empty rows
                                        if j != trailin_row:
                                            # checking whether date tuple is in the value or not
                                            if "(d" not in test_df.loc[j, k]:
                                                # Getting Dosage
                                                dosage_string = str(test_df.loc[j, k]).split('\n')[0]
                                                treatment_dict.setdefault("Dosage", []).append(dosage_string)
                                                # # Getting Procedure
                                                # procedure_string = str(test_df.loc[j, k]).replace('\n', ' ').replace(
                                                #     dosage_string, '', 1)
                                                # treatment_dict.setdefault("Procedure", []).append(procedure_string)

                        # Constructing dataframe from dictionary values
                        df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in treatment_dict.items()]))
                        # Concatinating dataframe to final_df and reassigning it to add multiple dataframes into one dataframe
                        final_df = pd.concat([final_df, df], axis=0)
                # Catch the exception
                except Exception as e:
                    print("Couldn't find more tables")
    # Filling NAN values with zeros
    final_df = final_df.fillna(0)
    # Reseting index to get new index
    final_df.reset_index(drop=True, inplace=True)
    # Dropping rows from dataframe based on given condition
    final_df = final_df.drop(final_df[(final_df.Treatment == 0) & (final_df.Dosage == 0)].index)
    # Replacing zero values to nan
    final_df.replace(0, np.nan, inplace=True)
    # Exporting dataframe to csv
    final_df.to_csv('./All_CSVs/' + today_date + '.csv', index=False)

    return "Dataframe cleaned and Exported in CSV format"


if __name__ == '__main__':
    # Calling function
    pre_process_data()