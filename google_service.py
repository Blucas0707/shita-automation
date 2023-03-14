from typing import List
from datetime import datetime, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Replace with the path to your service account credentials JSON file
SERVICE_ACCOUNT_FILE = 'google_credential.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/calendar']


class GoogleService:
    creds = creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES
    )

    def fetch_sheet_results(self, spread_sheet_id: str, sheet_name: str = None) -> List[list]:
        service = build('sheets', 'v4', credentials=self.creds)

        # Fetch the data from the sheet
        result = service.spreadsheets().values().get(
            spreadsheetId=spread_sheet_id,
            range=sheet_name
        ).execute()

        return [row for row in result['values'] if row]

    def batch_upsert_in_calendar(self, calendar_id: str, events: List[dict]):
        service = build('calendar', 'v3', credentials=self.creds)
        try:
            batch = service.new_batch_http_request()
            
            for event in events:
                batch.add(service.events().insert(calendarId=calendar_id, body=event))
                
            batch.execute()
        except HttpError as e:
            print(f'{str(e)} with event: {event}')
        
        print(f'{len(events)} events created: {event.get("htmlLink")}')
    
    def clean_calendar(self, calendar_id: str, to_clear_all: bool = False):
        service = build('calendar', 'v3', credentials=self.creds)
        now = datetime.utcnow()
        now_before_one_year = now - timedelta(days=365)
        now_before_one_year = now_before_one_year.isoformat() + 'Z'  # 'Z' indicates UTC time
        to_continue = True
        while to_continue:
            events_result = service.events().list(
                calendarId=calendar_id,
                timeMin=now_before_one_year,
                maxResults=100,
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            batch = service.new_batch_http_request()
            
            for event in events:
                if to_clear_all:
                    batch.add(service.events().delete(calendarId=calendar_id, eventId=event['id']))
                else:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    start_time = datetime.fromisoformat(start.replace('Z', '+00:00')).strftime('%Y-%m-%dT%H:%M:%S.%f')
                    if datetime.strptime(start_time, '%Y-%m-%dT%H:%M:%S.%f') < datetime.now():
                        batch.add(service.events().delete(calendarId=calendar_id, eventId=event['id']))
                    
            batch.execute()
            print(f'{len(events)} events deleted')
            
            to_continue = bool(events)
