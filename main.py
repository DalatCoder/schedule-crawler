# API
# https://qlgd.dlu.edu.vn/public/DrawingProfessorSchedule?YearStudy=2024-2025&TermID=HK02&Week=3&ProfessorID=011.031.00125&t=0.780616904287754

import requests
import json
from typing import Dict
from datetime import datetime
from bs4 import BeautifulSoup
from dataclasses import asdict, dataclass
from typing import Dict, List, Optional
import unicodedata

BASE_URL = "https://qlgd.dlu.edu.vn/public/DrawingProfessorSchedule"

@dataclass
class ClassSession:
    subject: str
    class_code: str
    class_name: str
    period: str  # keep original for reference
    period_begin: int
    period_end: int
    taught_lessons: str
    room: str
    content: str

@dataclass
class DaySchedule:
    morning: List[ClassSession]
    afternoon: List[ClassSession]
    evening: List[ClassSession]

@dataclass
class WeekInfo:
    week_number: int
    start_date: str
    end_date: str
    professor_name: str

class ScheduleCrawler:
    def __init__(self):
        self.year_study = "2024-2025"
        self.term_id = "HK02"
        self.professor_id = "011.031.00125"
    
    def build_url(self, week: int) -> str:
        timestamp = datetime.now().timestamp()
        return f"{BASE_URL}?YearStudy={self.year_study}&TermID={self.term_id}&Week={week}&ProfessorID={self.professor_id}&t={timestamp}"
    
    def _parse_period(self, period_str: str) -> tuple[int, int]:
        """Extract begin and end periods from period string like '1->4'"""
        try:
            if '->' in period_str:
                begin, end = period_str.split('->')
                return int(begin), int(end)
            return 0, 0
        except:
            return 0, 0

    def _parse_class_cell(self, cell) -> Optional[ClassSession]:
        if not cell.find('span'):
            return None
            
        spans = cell.find_all('span')
        if len(spans) < 7:  # Check if we have all required spans
            return None
            
        try:
            period_str = spans[3].text.replace('-Tiết:', '').strip() if spans[3] else ""
            period_begin, period_end = self._parse_period(period_str)
            
            return ClassSession(
                subject=spans[0].text.strip() if spans[0] else "",
                class_code=spans[1].text.replace('-Mã LHP:', '').strip() if spans[1] else "",
                class_name=spans[2].text.replace('-Lớp:', '').strip() if spans[2] else "",
                period=period_str,
                period_begin=period_begin,
                period_end=period_end,
                taught_lessons=spans[4].text.replace('-Đã dạy:', '').strip() if spans[4] else "",
                room=spans[5].text.replace('-Phòng :', '').strip() if spans[5] else "",
                content=spans[6].text.replace('-Nội dung :', '').strip() if spans[6] else ""
            )
        except Exception as e:
            print(f"Error parsing cell: {e}")
            return None

    def parse_schedule(self, html_content: str) -> Dict[str, DaySchedule]:
        soup = BeautifulSoup(html_content, 'html.parser')
        schedule = {}
        
        rows = soup.find('table').find_all('tr')[1:]  # Skip header row
        days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ nhật']
        
        for row, day in zip(rows, days):
            cells = row.find_all('td')
            schedule[day] = DaySchedule(
                morning=[self._parse_class_cell(cells[0])] if self._parse_class_cell(cells[0]) else [],
                afternoon=[self._parse_class_cell(cells[1])] if self._parse_class_cell(cells[1]) else [],
                evening=[self._parse_class_cell(cells[2])] if self._parse_class_cell(cells[2]) else []
            )
        
        return schedule

    def _normalize_text(self, text: str) -> str:
        """Normalize Vietnamese text by removing combining diacritics"""
        return ''.join(c for c in unicodedata.normalize('NFKD', text)
                      if not unicodedata.combining(c))

    def _extract_week_info(self, html_content: str) -> WeekInfo:
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            header_div = soup.find('div', style='font-weight:bold')
            spans = header_div.find_all('span')
            
            # Get raw texts
            week_span = spans[0].text.strip()
            professor_span = spans[1].text.strip()
            
            print(f"Raw week span: {week_span}")
            
            # Extract week number from original text
            week_number = int(''.join(filter(str.isdigit, week_span.split(':')[0])))
            
            # Extract dates using regex
            import re
            dates = re.findall(r'\d{2}/\d{2}/\d{4}', week_span)
            if len(dates) >= 2:
                start_date = dates[0]
                end_date = dates[1]
            else:
                raise ValueError(f"Could not find dates in text: {week_span}")
            
            # Extract professor name
            professor = professor_span.replace('Thời khóa biểu giảng viên:', '').strip()
            
            result = WeekInfo(
                week_number=week_number,
                start_date=start_date,
                end_date=end_date,
                professor_name=professor
            )
            print(f"Successfully parsed: {result}")
            return result
            
        except Exception as e:
            print(f"Error parsing header: {str(e)}")
            return WeekInfo(
                week_number=0,
                start_date="Unknown",
                end_date="Unknown",
                professor_name="Unknown"
            )

    def to_json_structure(self, schedule: Dict[str, DaySchedule], html_content: str) -> dict:
        week_info = self._extract_week_info(html_content)
        print(week_info)
        
        return {
            "metadata": asdict(week_info),
            "schedule": {
                day: {
                    "morning": [asdict(session) for session in day_schedule.morning],
                    "afternoon": [asdict(session) for session in day_schedule.afternoon],
                    "evening": [asdict(session) for session in day_schedule.evening]
                }
                for day, day_schedule in schedule.items()
            }
        }

    def fetch_schedule(self, week: int) -> Dict:
        url = self.build_url(week)
        try:
            response = requests.get(url, verify=False)
            requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
            
            if response.status_code == 200:
                schedule = self.parse_schedule(response.text)
                return self.to_json_structure(schedule, response.text)
            raise Exception(f"Failed to fetch schedule: {response.status_code}")
        except requests.exceptions.RequestException as req_err:
            print(f"Request Error occurred: {req_err}")
            raise

def main():
    crawler = ScheduleCrawler()
    schedule = crawler.fetch_schedule(3)
    
    # Generate filename with current date
    current_date = datetime.now().strftime("%Y%m%d")
    filename = f'schedule_{current_date}.json'
    
    # Save to JSON file with pretty printing
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(schedule, f, ensure_ascii=False, indent=2)
    
    print(f"Schedule has been saved to {filename}")

if __name__ == "__main__":
    main()

