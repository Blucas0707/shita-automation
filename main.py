import os
from typing import List
from datetime import datetime, date

from dotenv import load_dotenv

from google_service import GoogleService

load_dotenv()

SPREADSHEET_ID = os.environ.get('SHITA_GOOGLE_SHEET_ID')

SHEET_NAME = os.environ.get('SHITA_GOOGLE_SHEET_NAME')

CALENDAR_ID = os.environ.get('SHITA_CALENDAR_ID')


CONFIG_D = {
    'location': '104台北市中山區林森北路138巷12號B1',
    'timezone': 'Asia/Taipei',
}


def _format_start_end_datetime(date_description: str, time_description: str) -> str:
    
    event_month, event_date = _parse_month_date(date_description)
    start_hour, start_min = _parse_hour_min(time_description.split('-')[0])
    end_hour, end_min = _parse_hour_min(time_description.split('-')[1])
    
    today = date.today()
    year = today.year
    
    if event_month < today.month:
        year += 1
    
    start_dt = datetime(year, event_month, event_date, start_hour, start_min).strftime('%Y-%m-%dT%H:%M:%S')
    
    event_date = event_date + 1 if end_hour < start_hour else event_date
    end_dt = datetime(year, event_month, event_date, end_hour, end_min).strftime('%Y-%m-%dT%H:%M:%S')

    return start_dt, end_dt


def _parse_month_date(time_description: str):
    '''parse month and date in a description
        example: '6月10日 週六'
    '''
    month = time_description.split('月')[0]
    date = time_description.split('月')[1].split('日')[0]
    return int(month), int(date)


def _parse_hour_min(time_description: str):
    '''parse hour and min in a description
        example: '1600'
    '''
    return int(time_description[:2]), int(time_description[2:])


def format_result_to_event_ds(results: List[list]) -> List[dict]:
    '''format data in sheet to calendar event
        example: ['6月10日 週六', '1600-2200', 'FridaySwing']
    '''
    event_ds = []
    for result in results:
        start_dt, end_dt = _format_start_end_datetime(result[0], result[1])
        
        event_ds.append(
            {
                'summary': result[2],
                'location': CONFIG_D['location'],
                'description': f'Occupied by {result[2]}',
                'start': {
                    'dateTime': start_dt,
                    'timeZone': CONFIG_D['timezone'],
                },
                'end': {
                    'dateTime': end_dt,
                    'timeZone': CONFIG_D['timezone'],
                },
                'reminders': {
                    'useDefault': True,
                },
            }
        )
    
    return event_ds


def execute():
    
    google_service = GoogleService()
    
    google_service.clean_calendar(CALENDAR_ID, to_clear_all=True)
    
    results = google_service.fetch_sheet_results(SPREADSHEET_ID, SHEET_NAME)
    if not results:
        return

    event_ds = format_result_to_event_ds(results=results[1:])

    google_service.batch_upsert_in_calendar(CALENDAR_ID, event_ds)


if __name__ == '__main__':
    execute()
