import re
import json
import requests
from bs4 import BeautifulSoup
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from typing import Dict, List, Optional

BASE_URL = "https://qlgd.dlu.edu.vn/public/DrawingClassStudentSchedules_Mau2"

@dataclass
class StudentSession:
    subject: str
    class_code: str
    class_name: str
    period: str
    period_begin: int
    period_end: int
    time_begin: str
    time_end: str
    room: str
    teacher_name: str

    def to_dict(self):
        """Convert session to dictionary for JSON serialization"""
        return {
            'subject': self.subject,
            'class_code': self.class_code,
            'class_name': self.class_name,
            'period': self.period,
            'period_begin': self.period_begin,
            'period_end': self.period_end,
            'time_begin': self.time_begin,
            'time_end': self.time_end,
            'room': self.room,
            'teacher_name': self.teacher_name
        }

@dataclass
class StudentDaySchedule:
    morning: List[StudentSession]
    afternoon: List[StudentSession]
    evening: List[StudentSession]

class StudentScheduleCrawler:
    def __init__(self):
        self.year_study = "2024-2025"
        self.term_id = "HK02"
        self.class_id = "KTK48A"  # Example class ID
        self.period_map = self._initialize_period_map()

    def build_url(self, week: int) -> str:
        timestamp = datetime.now().timestamp()
        return f"{BASE_URL}?YearStudy={self.year_study}&TermID={self.term_id}&Week={week}&ClassStudentID={self.class_id}&t={timestamp}"

    def _initialize_period_map(self) -> Dict[int, tuple]:
        """Initialize mapping of period numbers to actual times"""
        periods = {}
        
        # Morning periods (1-6)
        current_time = datetime.strptime("07:30", "%H:%M")  # Start at 7:30
        for period in range(1, 7):
            start_time = current_time.strftime("%H:%M")
            end_time = (current_time + timedelta(minutes=45)).strftime("%H:%M")
            periods[period] = (start_time, end_time)
            
            # Add break time
            current_time += timedelta(minutes=45)  # Add period duration
            if period == 3:  # Long break after period 3
                current_time += timedelta(minutes=20)  # Long break
            else:
                current_time += timedelta(minutes=5)  # Short break
        
        # Afternoon periods (7-10)
        current_time = datetime.strptime("13:00", "%H:%M")  # Start at 13:00
        for period in range(7, 11):
            start_time = current_time.strftime("%H:%M")
            end_time = (current_time + timedelta(minutes=45)).strftime("%H:%M")
            periods[period] = (start_time, end_time)
            
            # Add break time
            current_time += timedelta(minutes=45)  # Add period duration
            if period == 8:  # Long break after period 8
                current_time += timedelta(minutes=10)  # Long break
            else:
                current_time += timedelta(minutes=5)  # Short break
        
        # Evening periods (11-14)
        current_time = datetime.strptime("16:40", "%H:%M")  # Start at 16:40
        for period in range(11, 15):
            start_time = current_time.strftime("%H:%M")
            end_time = (current_time + timedelta(minutes=45)).strftime("%H:%M")
            periods[period] = (start_time, end_time)
            
            # Only add short breaks between periods
            current_time += timedelta(minutes=50)  # 45 mins + 5 mins break
        
        return periods

    def _parse_period(self, period_str: str) -> tuple[int, int]:
        """Extract begin and end periods from period string like '- Tiết: 1-2'"""
        try:
            # Remove "- Tiết:" prefix and whitespace
            cleaned = period_str.replace('- Tiết:', '').strip()
            if '-' in cleaned:
                begin, end = cleaned.split('-')
                return int(begin), int(end)
            return 0, 0
        except:
            return 0, 0

    def _parse_class_cell(self, cell) -> List[Optional[StudentSession]]:
        if not cell.find('span'):
            return []

        # Split content by <hr> tag if exists
        sessions = []
        subject_blocks = cell.find_all(['hr', 'span'])
        
        current_session = []
        for block in subject_blocks:
            if block.name == 'hr':
                if current_session:
                    session = self._create_session_from_spans(current_session)
                    if session:
                        sessions.append(session)
                current_session = []
            else:
                current_session.append(block)
                
        # Don't forget the last session
        if current_session:
            session = self._create_session_from_spans(current_session)
            if session:
                sessions.append(session)
                
        return sessions if sessions else []

    def _create_session_from_spans(self, spans) -> Optional[StudentSession]:
        if len(spans) < 6:
            return None

        try:
            # Extract period string with "Tiết" prefix
            period_str = spans[3].text.strip() if spans[3] else ""
            period_begin, period_end = self._parse_period(period_str)

            # Get time slots for the periods
            time_begin = self.period_map[period_begin][0] if period_begin in self.period_map else "00:00"
            time_end = self.period_map[period_end][1] if period_end in self.period_map else "00:00"

            return StudentSession(
                subject=spans[0].text.strip(),
                class_code=spans[1].text.replace('- Nhóm:', '').strip(),
                class_name=spans[2].text.replace('- Lớp:', '').strip(),
                period=period_str,
                period_begin=period_begin,
                period_end=period_end,
                time_begin=time_begin,
                time_end=time_end,
                room=spans[4].text.replace('- Phòng:', '').strip(),
                teacher_name=spans[5].text.replace('- GV:', '').strip()
            )
        except Exception as e:
            print(f"Error parsing spans: {e}")
            return None

    def fetch_schedule(self, week: int) -> Dict:
        url = self.build_url(week)
        response = requests.get(url, verify=False)
        requests.packages.urllib3.disable_warnings()

        if response.status_code != 200:
            raise Exception(f"Failed to fetch schedule: {response.status_code}")

        schedule = self.parse_schedule(response.text)
        metadata = self._extract_metadata(response.text)
        
        return {
            "metadata": metadata,
            "schedule": schedule
        }

    def _extract_metadata(self, html_content: str) -> Dict:
        soup = BeautifulSoup(html_content, 'html.parser')
        header_div = soup.find('div', style='font-weight:bold')
        spans = header_div.find_all('span')

        week_info = spans[0].text.strip()
        class_info = spans[1].text.strip()

        # Extract week number and dates
        week_number = int(''.join(filter(str.isdigit, week_info.split(':')[0])))
        dates = re.findall(r'\d{2}/\d{2}/\d{4}', week_info)
        
        return {
            "week_number": week_number,
        # Find the schedule table
            "start_date": dates[0] if len(dates) > 0 else "",
            "end_date": dates[1] if len(dates) > 1 else "",
            "class_name": class_info.replace("Thời khóa biểu lớp:", "").strip()
        }

    def parse_schedule(self, html_content: str) -> Dict[str, Dict]:
        """Parse the HTML content and extract schedule data"""
        soup = BeautifulSoup(html_content, 'html.parser')
        schedule = {}
        
        # Find the schedule table
        table = soup.find('table')
        if not table:
            raise Exception("Schedule table not found in HTML content")
            
        # Get all rows except header
        rows = table.find_all('tr')[1:]  # Skip header row
        days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ nhật']
        
        # Process each row (day)
        for row, day in zip(rows, days):
            cells = row.find_all('td')
            if len(cells) >= 3:  # Should have morning, afternoon, evening cells
                morning_sessions = self._parse_class_cell(cells[0])
                afternoon_sessions = self._parse_class_cell(cells[1])
                evening_sessions = self._parse_class_cell(cells[2])
                
                schedule[day] = {
                    'morning': [s.to_dict() for s in morning_sessions] if morning_sessions else [],
                    'afternoon': [s.to_dict() for s in afternoon_sessions] if afternoon_sessions else [],
                    'evening': [s.to_dict() for s in evening_sessions] if evening_sessions else []
                }
        
        return schedule

def main():
    crawler = StudentScheduleCrawler()
    schedule = crawler.fetch_schedule(3)  # Fetch week 3
    
    # Save to JSON file
    filename = f'student_schedule_{datetime.now().strftime("%Y%m%d")}.json'
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(schedule, f, ensure_ascii=False, indent=2)

    print(f"Schedule saved to {filename}")

if __name__ == "__main__":
    main()