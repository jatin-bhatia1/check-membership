import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import re

def GenerateExcelFile(Participants_Data, file_name):

    print("-------- Inserting Generated Data into Excel File ---------------")

    # Collect the date-time info
    datetimestamp = str(datetime.now(ZoneInfo('Europe/Paris')))
    datetimestamp_list = datetimestamp.split(" ")
    date = datetimestamp_list[0]
    time = (datetimestamp_list[1].split(".")[0]).replace(':', '-')

    # Convert the list of data into Dataframes
    df = pd.DataFrame(Participants_Data)
    
    # format the file name correctly
    full_file_name = ("{0}-D-{1}-T-{2}.xlsx").format(file_name, date, time)

    # Generate the excel file with data
    df.to_excel(full_file_name, sheet_name=file_name)

    print(("A new file {0} with data has been generated").format(full_file_name))

    return full_file_name

