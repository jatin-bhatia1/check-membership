import pandas as pd
import os

def GenerateExcelFile(Participants_Data, file_name):

    print("-------- Inserting Generated Data into Excel File ---------------")

    # Convert the list of data into Dataframes
    df = pd.DataFrame(Participants_Data)
    
    # format the file name correctly
    full_file_name = ("{0}.xlsx").format(file_name)

    # Generate the excel file with data and sheet name can't exceed more than 31 characters
    df.to_excel(full_file_name, sheet_name=(file_name[:30]))

    print(("A new file {0} with data has been generated").format(full_file_name))

    return full_file_name

