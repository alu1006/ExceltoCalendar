from __future__ import print_function
from asyncio import events
from calendar import calendar

from datetime import datetime, timezone, timedelta
import os.path
import re

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import ExceltoCalendar

ExceltoCalendar.CalendarFormat()


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']
creds = None
def connection():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        service = build('calendar', 'v3', credentials=creds)
        return service
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
def show_calendar(service):
    page_token=None
    calendar_list = service.calendarList().list(pageToken=page_token).execute()
    # print(calendar_list)
    while True:
        for calendar_list_entry in calendar_list['items']:
            print(calendar_list_entry['summary'],calendar_list_entry['id'])
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
        
def show_calendar_events(service,calendar_id):
    page_token = None
    while True:
        events = service.events().list(calendarId=calendar_id, pageToken=page_token).execute()
        for event in events['items']:
            print(event)
        page_token = events.get('nextPageToken')
        if not page_token:
            break

def create_calendar(service,summary,timezone='Asia/Taipei'):
    calendar ={
        'summary':summary,
        'timezone':timezone
    }
    response = service.calendars().insert(body=calendar).execute()
    print(response['id'])
    
def create_calendar_event(service,calendarId,summary,start_time,end_time,description=None):
    try:
        
        

        # Call the Calendar API
        # 設定為 +8 時區
        tz = timezone(timedelta(hours=+8))

        # 取得現在時間、指定時區、轉為 ISO 格式
        now=datetime.now(tz).isoformat(timespec='seconds')
        
        template={
            "summary": summary,
            
            "description": description,
            "start": {
                "dateTime": start_time,
                "timeZone": "Asia/Taipei"
            },
            "end": {
                "dateTime": end_time,
                "timeZone": "Asia/Taipei"
            },
            
            
        }

        response = service.events().insert(calendarId=calendarId, body=template).execute()
        # print(response['summary'])
        
        #show calendar list
        '''
        
        '''

    except HttpError as error:
        print('An error occurred: %s' % error)


def delete_calendar_event(service,calendar_id):
    page_token = None
    while True:
        events = service.events().list(calendarId=calendar_id, pageToken=page_token).execute()
        for event in events['items']:
            if event['start']['date']<='2022-07-06':
                
                service.events().delete(calendarId=calendar_id,eventId=event['id']).execute()
        page_token = events.get('nextPageToken')
        if not page_token:
            break
