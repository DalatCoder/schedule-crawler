from datetime import datetime, timedelta
import json
import uuid

class ICSExporter:
    def __init__(self):
        self.timezone = "Asia/Ho_Chi_Minh"
        
    def _generate_uid(self):
        return str(uuid.uuid4()).replace("-", "")
        
    def _format_datetime(self, date_str: str, time_str: str) -> str:
        """Convert date and time to iCal format"""
        # Convert DD/MM/YYYY to YYYYMMDD
        date_parts = date_str.split('/')
        formatted_date = f"{date_parts[2]}{date_parts[1]}{date_parts[0]}"
        # Convert HH:MM to HHMMSS
        formatted_time = time_str.replace(":", "") + "00"
        return f"{formatted_date}T{formatted_time}"

    def create_ics_content_from_data(self, schedule_data):
        """Create ICS content directly from schedule data dictionary"""
        try:
            # Use the schedule data directly instead of loading from file
            return self._generate_ics_content(schedule_data)
        except Exception as e:
            raise Exception(f"Failed to create ICS content: {str(e)}")

    def create_ics_content(self, schedule_file):
        """Create ICS content from schedule JSON file"""
        try:
            with open(schedule_file, 'r', encoding='utf-8') as f:
                schedule_data = json.load(f)
            return self._generate_ics_content(schedule_data)
        except Exception as e:
            raise Exception(f"Failed to create ICS content: {str(e)}")

    def _generate_ics_content(self, schedule_data):
        """Common method to generate ICS content from schedule data"""
        metadata = schedule_data['metadata']
        schedule = schedule_data['schedule']
        
        # Start building ICS content
        ics_lines = [
            "BEGIN:VCALENDAR",
            "PRODID:-//DalatCoder//Schedule Exporter//EN",
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
        
        # Create events for each session
        for day, sessions in schedule.items():
            # Calculate the actual date for this day
            current_date = base_date + timedelta(days=day_map[day])
            current_date_str = current_date.strftime("%d/%m/%Y")
            
            for period in ['morning', 'afternoon', 'evening']:
                for session in sessions[period]:
                    if not session:  # Skip empty sessions
                        continue
                        
                    start_dt = self._format_datetime(current_date_str, session['time_begin'])
                    end_dt = self._format_datetime(current_date_str, session['time_end'])
                    
                    description = (
                        f"Mã lớp: {session['class_code']}\\n"
                        f"Lớp: {session['class_name']}\\n"
                        f"Tiết: {session['period']}\\n"
                        f"Đã dạy: {session['taught_lessons']}"
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
        
        # Close the calendar
        ics_lines.append("END:VCALENDAR")
        
        return "\r\n".join(ics_lines)

def main():
    exporter = ICSExporter()
    
    # Find the latest schedule file
    import os
    schedule_files = [f for f in os.listdir('.') if f.startswith('schedule_') and f.endswith('.json')]
    if not schedule_files:
        print("No schedule files found!")
        return
        
    latest_file = max(schedule_files)
    print(f"Exporting schedule from {latest_file}")
    
    # Generate ICS content
    ics_content = exporter.create_ics_content(latest_file)
    
    # Save to file
    output_file = f"teaching_schedule_{datetime.now().strftime('%Y%m%d')}.ics"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(ics_content)
    
    print(f"Schedule has been exported to {output_file}")

if __name__ == "__main__":
    main()
