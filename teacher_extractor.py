from bs4 import BeautifulSoup
import json
from dataclasses import dataclass, asdict
import re

@dataclass
class Teacher:
    id: str
    last_name: str
    first_name: str
    full_name: str

class TeacherExtractor:
    def extract_name_parts(self, full_name: str) -> tuple[str, str]:
        """Extract first and last name from full name format: 'Last, First Middle'"""
        parts = full_name.split(',', 1)
        if len(parts) == 2:
            last_name = parts[0].strip()
            first_name = parts[1].strip()
            return last_name, first_name
        return full_name, ""

    def parse_teachers(self, html_file: str) -> list[Teacher]:
        with open(html_file, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f.read(), 'html.parser')
        
        teachers = []
        for option in soup.find_all('option'):
            teacher_id = option['value']
            full_name = option.text.strip()
            last_name, first_name = self.extract_name_parts(full_name)
            
            teacher = Teacher(
                id=teacher_id,
                last_name=last_name,
                first_name=first_name,
                full_name=full_name
            )
            teachers.append(teacher)
        
        return teachers

    def save_to_json(self, teachers: list[Teacher], output_file: str):
        data = {
            "count": len(teachers),
            "teachers": [asdict(t) for t in teachers]
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

def main():
    extractor = TeacherExtractor()
    teachers = extractor.parse_teachers('teachers.html')
    extractor.save_to_json(teachers, 'teachers.json')
    print(f"Extracted {len(teachers)} teachers to teachers.json")

if __name__ == "__main__":
    main()
