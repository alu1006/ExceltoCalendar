from __future__ import print_function
from asyncio import events
from calendar import calendar

from datetime import datetime, timezone, timedelta
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
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

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        # 設定為 +8 時區
        tz = timezone(timedelta(hours=+8))

        # 取得現在時間、指定時區、轉為 ISO 格式
        now=datetime.now(tz).isoformat(timespec='seconds')
        
        template={
            "summary": "Important event!",
            "location": "Virtual event (Slack)",
            "description": "This a description",
            "start": {
                "dateTime": now,
                "timeZone": "Asia/Taipei"
            },
            "end": {
                "dateTime": now,
                "timeZone": "Asia/Taipei"
            },
            "
            
        }

        response = service.events().insert(calendarId="c_fclsqdl2se8oalno5pn54n7604@group.calendar.google.com", body=template).execute()
        print(response['summary'])
        
        page_token=None
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        print(calendar_list)
        while True:
            calendar_list = service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print(calendar_list_entry['summary'])
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break


    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()