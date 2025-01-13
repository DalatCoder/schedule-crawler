import os
from datetime import datetime, timedelta
from ics_exporter import ICSExporter

class StudentICSExporter(ICSExporter):
    def _generate_ics_content(self, schedule_data):
        """Override to handle student-specific schedule format"""
        metadata = schedule_data['metadata']
        schedule = schedule_data['schedule']
        
        ics_lines = [
            "BEGIN:VCALENDAR",
            "PRODID:-//DalatCoder//Student Schedule Exporter//EN",
            "VERSION:2.0",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-TIMEZONE:{self.timezone}"
        ]
        
        # Map Vietnamese days to their dates
        day_map = {
            'Thứ 2': 0, 'Thứ 3': 1, 'Thứ 4': 2,
            'Thứ 5': 3, 'Thứ 6': 4, 'Thứ 7': 5,
            'Chủ nhật': 6
        }
        
        base_date = datetime.strptime(metadata['start_date'], "%d/%m/%Y")
        
        for day, sessions in schedule.items():
            current_date = base_date + timedelta(days=day_map[day])
            current_date_str = current_date.strftime("%d/%m/%Y")
            
            for period in ['morning', 'afternoon', 'evening']:
                for session in sessions[period]:
                    if not session:
                        continue
                        
                    start_dt = self._format_datetime(current_date_str, session['time_begin'])
                    end_dt = self._format_datetime(current_date_str, session['time_end'])
                    
                    description = (
                        f"Mã lớp: {session['class_code']}\\n"
                        f"Lớp: {session['class_name']}\\n"
                        f"Tiết: {session['period']}\\n"
                        f"Giảng viên: {session['teacher_name']}"
                    )
                    
                    event_lines = [
                        "BEGIN:VEVENT",
                        f"UID:{self._generate_uid()}",
                        f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                        f"DTSTART;TZID={self.timezone}:{start_dt}",
                        f"DTEND;TZID={self.timezone}:{end_dt}",
                        f"SUMMARY:{session['subject']}",
                        f"LOCATION:{session['room']}",
                        f"DESCRIPTION:{description}",
                        "STATUS:CONFIRMED",
                        "SEQUENCE:0",
                        "END:VEVENT"
                    ]
                    
                    ics_lines.extend(event_lines)
        
        ics_lines.append("END:VCALENDAR")
        return "\r\n".join(ics_lines)

def main():
    # Find the latest student schedule file
    schedule_files = [f for f in os.listdir('.') if f.startswith('student_schedule_') and f.endswith('.json')]
    if not schedule_files:
        print("No student schedule files found!")
        return
        
    latest_file = max(schedule_files)
    print(f"Exporting schedule from {latest_file}")
    
    exporter = StudentICSExporter()
    ics_content = exporter.create_ics_content(latest_file)
    
    output_file = f"student_schedule_{datetime.now().strftime('%Y%m%d')}.ics"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(ics_content)
    
    print(f"Schedule has been exported to {output_file}")

if __name__ == "__main__":
    main()
