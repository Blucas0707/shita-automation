import os
import json
from typing import List
from datetime import datetime, timedelta

from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()


# Replace with the path to your service account credentials JSON file
SERVICE_ACCOUNT_INO_DICT = json.loads(os.getenv('GOOGLE_CREDENTIAL_JSON'))

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/calendar']


class GoogleService:
    creds = service_account.Credentials.from_service_account_info(
        SERVICE_ACCOUNT_INO_DICT,
        scopes=SCOPES
    )
    
    calendar_service = build('calendar', 'v3', credentials=creds, static_discovery=False)
    
    sheet_service = build('sheets', 'v4', credentials=creds, static_discovery=False)

    def fetch_sheet_results(self, spread_sheet_id: str, sheet_name: str = None) -> List[list]:
        # Fetch the data from the sheet
        result = self.sheet_service.spreadsheets().values().get(
            spreadsheetId=spread_sheet_id,
            range=sheet_name
        ).execute()

        return [row for row in result['values'] if row]

    def batch_upsert_in_calendar(self, calendar_id: str, events: List[dict]):
        try:
            batch = self.calendar_service.new_batch_http_request()
            
            for event in events:
                batch.add(self.calendar_service.events().insert(calendarId=calendar_id, body=event))
                
            batch.execute()
        except HttpError as e:
            print(f'{str(e)} with event: {event}')
        
        print(f'{len(events)} events created: {event.get("htmlLink")}')
    
    def clean_calendar(self, calendar_id: str, to_clear_all: bool = False):
        now = datetime.utcnow()
        now_before_one_year = now - timedelta(days=365)
        now_before_one_year = now_before_one_year.isoformat() + 'Z'  # 'Z' indicates UTC time
        
        to_continue = True
        while to_continue:
            events_result = self.calendar_service.events().list(
                calendarId=calendar_id,
                timeMin=now_before_one_year,
                maxResults=100,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            batch = self.calendar_service.new_batch_http_request()
            
            for event in events:
                if to_clear_all:
                    batch.add(self.calendar_service.events().delete(calendarId=calendar_id, eventId=event['id']))
                else:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    start_time = datetime.fromisoformat(start.replace('Z', '+00:00')).strftime('%Y-%m-%dT%H:%M:%S.%f')
                    if datetime.strptime(start_time, '%Y-%m-%dT%H:%M:%S.%f') < datetime.now():
                        batch.add(self.calendar_service.events().delete(calendarId=calendar_id, eventId=event['id']))
                    
            batch.execute()
            print(f'{len(events)} events deleted')
            
            to_continue = bool(events)
