
import calendar
import pandas as pd
import ExceltoCalendar ,PandasReadCSV, quickstart
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta, timezone



def main():
    #read the data from csv or excel file
    filename='./Google Calendar/Schedule.xlsx'
    df=pd.read_excel(filename,sheet_name='google_calendar')
    
    #connect to Google Calendar
    service=quickstart.connection()
    
    #show the calendar
    quickstart.show_calendar(service)
    
    # quickstart.create_calendar(service,'班表')
    
    #create schedule 
    calendarId='c_oco79o9sd2qleeega5f7a5aefo@group.calendar.google.com'
    quickstart.show_calendar_events(service,calendarId)
    # 取得現在時間、指定時區、轉為 ISO 格式
    # tz = timezone(timedelta(hours=+8))
    # print(df.values[0])
    # now=datetime.now(tz).isoformat(timespec='seconds')
    # print(now)
    
    # for value in df.values:
    #     summary=value[0]
    #     st=value[1]
    #     et=value[3]
    #     if pd.isna(value[2]):
            
    #         st= datetime.strptime(st+' '+'0:0:0','%Y/%m/%d %H:%M:%S').isoformat()
    #         et= datetime.strptime(et+' '+'23:59:59','%Y/%m/%d %H:%M:%S').isoformat()
    #     else:
    #         st= datetime.strptime(st+' '+value[2],'%Y/%m/%d %H:%M').isoformat()
    #         et= datetime.strptime(et+' '+value[4],'%Y/%m/%d %H:%M').isoformat()
    #     des=value[6] if pd.isna(value[6]) else ''
    #     quickstart.create_calendar_event(service,calendarId,summary,st,et,des,)
    
   
        
    


if __name__ == '__main__':
    main()