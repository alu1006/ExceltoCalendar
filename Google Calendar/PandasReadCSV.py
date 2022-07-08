import pandas as pd
def read_data(filename):
    
    df=pd.read_excel(filename,sheet_name='google_calendar')
    # print(df.values)

    # for value in df.values:
    #     print(value)
        
# filename ='./Google Calender/Schedule.xlsx'
# read_data(filename)