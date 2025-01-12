from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import json
import os.path
import pickle

SCOPES = ['https://www.googleapis.com/auth/calendar']

class GoogleCalendarSync:
    def __init__(self):
        self.creds = self._get_credentials()
        self.service = build('calendar', 'v3', credentials=self.creds)
        
    def _get_credentials(self):
        creds = None
        # The file token.pickle stores the user's access and refresh tokens
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
                
        # If there are no (valid) credentials available, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return creds

    def create_event(self, subject, location, start_time, end_time, description):
        # Convert times to RFC3339 format
        date_str = datetime.now().strftime("%Y-%m-%d")
        start_datetime = f"{date_str}T{start_time}:00+07:00"
        end_datetime = f"{date_str}T{end_time}:00+07:00"
        
        event = {
            'summary': subject,
            'location': location,
            'description': description,
            'start': {
                'dateTime': start_datetime,
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'end': {
                'dateTime': end_datetime,
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'reminders': {
                'useDefault': True
            },
        }

        event = self.service.events().insert(calendarId='primary', body=event).execute()
        print(f'Event created: {event.get("htmlLink")}')
        return event

    def sync_schedule(self, schedule_file):
        with open(schedule_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        metadata = data['metadata']
        schedule = data['schedule']
        
        # Convert start_date from DD/MM/YYYY to YYYY-MM-DD
        start_parts = metadata['start_date'].split('/')
        schedule_date = f"{start_parts[2]}-{start_parts[1]}-{start_parts[0]}"
        
        # Map Vietnamese days to datetime weekday numbers
        day_map = {
            'Thứ 2': 0, 'Thứ 3': 1, 'Thứ 4': 2, 
            'Thứ 5': 3, 'Thứ 6': 4, 'Thứ 7': 5, 
            'Chủ nhật': 6
        }
        
        # Calculate dates for each day in the schedule
        base_date = datetime.strptime(schedule_date, "%Y-%m-%d")
        created_events = []
        
        for day, sessions in schedule.items():
            # Calculate the actual date for this day
            days_to_add = day_map[day]
            current_date = base_date + timedelta(days=days_to_add)
            date_str = current_date.strftime("%Y-%m-%d")
            
            for period in ['morning', 'afternoon', 'evening']:
                for session in sessions[period]:
                    description = (
                        f"Mã lớp: {session['class_code']}\n"
                        f"Lớp: {session['class_name']}\n"
                        f"Tiết: {session['period']}\n"
                        f"Đã dạy: {session['taught_lessons']}"
                    )
                    
                    event = self.create_event(
                        subject=session['subject'],
                        location=session['room'],
                        start_time=session['time_begin'],
                        end_time=session['time_end'],
                        description=description
                    )
                    created_events.append(event)
        
        return created_events

def main():
    calendar = GoogleCalendarSync()
    # Use the most recent schedule file
    schedule_files = [f for f in os.listdir('.') if f.startswith('schedule_') and f.endswith('.json')]
    if not schedule_files:
        print("No schedule files found!")
        return
        
    latest_file = max(schedule_files)
    print(f"Syncing schedule from {latest_file}")
    
    # events = calendar.sync_schedule(latest_file)
    # print(f"Successfully created {len(events)} calendar events")

if __name__ == "__main__":
    main()
