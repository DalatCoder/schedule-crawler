import sys
import json
from datetime import datetime
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QComboBox, QPushButton, 
    QTextEdit, QTabWidget, QFileDialog, QMessageBox,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from main import ScheduleCrawler
from ics_exporter import ICSExporter

class CrawlerWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, year_study, term_id, professor_id, week):
        super().__init__()
        self.year_study = year_study
        self.term_id = term_id
        self.professor_id = professor_id
        self.week = week

    def run(self):
        try:
            self.progress.emit("Initializing crawler...")
            crawler = ScheduleCrawler()
            crawler.year_study = self.year_study
            crawler.term_id = self.term_id
            crawler.professor_id = self.professor_id
            
            self.progress.emit("Fetching schedule...")
            schedule = crawler.fetch_schedule(self.week)
            self.progress.emit("Schedule fetched successfully!")
            self.finished.emit(schedule)
        except Exception as e:
            self.error.emit(str(e))

class ICSWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, schedule_file):
        super().__init__()
        self.schedule_file = schedule_file

    def run(self):
        try:
            self.progress.emit("Creating ICS file...")
            exporter = ICSExporter()
            ics_content = exporter.create_ics_content(self.schedule_file)
            self.progress.emit("ICS content generated successfully!")
            self.finished.emit(ics_content)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Schedule Manager")
        self.setMinimumSize(800, 600)
        
        # Create the main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Create tabs
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # Create Crawler tab
        crawler_tab = QWidget()
        crawler_layout = QVBoxLayout(crawler_tab)
        tabs.addTab(crawler_tab, "Schedule Crawler")
        
        # Create ICS Exporter tab
        ics_tab = QWidget()
        ics_layout = QVBoxLayout(ics_tab)
        tabs.addTab(ics_tab, "ICS Exporter")
        
        # Setup Crawler tab
        self.setup_crawler_tab(crawler_layout)
        
        # Setup ICS Exporter tab
        self.setup_ics_tab(ics_layout)
        
        # Load configuration data
        self.load_config_data()

    def load_config_data(self):
        try:
            # Load years
            with open('config/year_studies.json', 'r', encoding='utf-8') as f:
                years = json.load(f)
                self.year_combo.addItems([year['value'] for year in years])
            
            # Load terms
            with open('config/terms.json', 'r', encoding='utf-8') as f:
                terms = json.load(f)
                self.term_combo.addItems([term['value'] for term in terms])
            
            # Load teachers with modified storage
            with open('config/teachers.json', 'r', encoding='utf-8') as f:
                self.all_teachers = json.load(f)
                for teacher in self.all_teachers:
                    self.teacher_combo.addItem(
                        f"{teacher['full_name']} ({teacher['id']})",
                        teacher['id']
                    )
            
            # Load weeks
            with open('config/weeks.json', 'r', encoding='utf-8') as f:
                weeks = json.load(f)
                self.week_combo.addItems([str(week['label']) for week in weeks])
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load configuration: {str(e)}")

    def setup_crawler_tab(self, layout):
        # Create form layout for inputs
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        layout.addWidget(form_widget)
        
        # Today's date
        today_layout = QHBoxLayout()
        today_layout.addWidget(QLabel("Today:"))
        self.today_label = QLabel(datetime.now().strftime("%d/%m/%Y"))
        today_layout.addWidget(self.today_label)
        form_layout.addLayout(today_layout)
        
        # Schedule metadata section
        metadata_layout = QHBoxLayout()
        metadata_layout.addWidget(QLabel("Schedule Info:"))
        self.metadata_label = QLabel("No schedule loaded")
        self.metadata_label.setWordWrap(True)
        metadata_layout.addWidget(self.metadata_label)
        form_layout.addLayout(metadata_layout)
        
        # Year selection
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel("Academic Year:"))
        self.year_combo = QComboBox()
        year_layout.addWidget(self.year_combo)
        form_layout.addLayout(year_layout)
        
        # Term selection
        term_layout = QHBoxLayout()
        term_layout.addWidget(QLabel("Term:"))
        self.term_combo = QComboBox()
        term_layout.addWidget(self.term_combo)
        form_layout.addLayout(term_layout)
        
        # Teacher selection with filter
        teacher_layout = QVBoxLayout()  # Changed to VBox to stack filter and combo
        teacher_header = QHBoxLayout()
        teacher_header.addWidget(QLabel("Teacher:"))
        
        # Add search box
        self.teacher_filter = QLineEdit()
        self.teacher_filter.setPlaceholderText("Search teacher...")
        self.teacher_filter.textChanged.connect(self.filter_teachers)
        teacher_header.addWidget(self.teacher_filter)
        
        teacher_layout.addLayout(teacher_header)
        
        self.teacher_combo = QComboBox()
        self.teacher_combo.setMaxVisibleItems(10)  # Show more items in dropdown
        teacher_layout.addWidget(self.teacher_combo)
        form_layout.addLayout(teacher_layout)
        
        # Store original teacher items for filtering
        self.all_teachers = []
        
        # Week selection
        week_layout = QHBoxLayout()
        week_layout.addWidget(QLabel("Week:"))
        self.week_combo = QComboBox()
        week_layout.addWidget(self.week_combo)
        form_layout.addLayout(week_layout)
        
        # Fetch button
        self.fetch_button = QPushButton("Fetch Schedule")
        self.fetch_button.clicked.connect(self.fetch_schedule)
        form_layout.addWidget(self.fetch_button)
        
        # Progress bar
        self.crawler_progress = QProgressBar()
        self.crawler_progress.setTextVisible(True)
        self.crawler_progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.crawler_progress)
        
        # Schedule Table
        self.schedule_table = QTableWidget()
        self.schedule_table.setColumnCount(4)
        self.schedule_table.setHorizontalHeaderLabels(['Time', 'Subject', 'Room', 'Content'])
        self.schedule_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.schedule_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.schedule_table)
        
        # Results area (make it smaller since we have the table now)
        self.crawler_output = QTextEdit()
        self.crawler_output.setReadOnly(True)
        self.crawler_output.setMaximumHeight(100)  # Limit height
        layout.addWidget(self.crawler_output)

    def setup_ics_tab(self, layout):
        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        file_layout.addWidget(self.file_label)
        
        select_file_btn = QPushButton("Select Schedule File")
        select_file_btn.clicked.connect(self.select_schedule_file)
        file_layout.addWidget(select_file_btn)
        layout.addLayout(file_layout)
        
        # Export button
        self.export_button = QPushButton("Export to ICS")
        self.export_button.clicked.connect(self.export_to_ics)
        self.export_button.setEnabled(False)
        layout.addWidget(self.export_button)
        
        # Progress bar
        self.ics_progress = QProgressBar()
        self.ics_progress.setTextVisible(True)
        self.ics_progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.ics_progress)
        
        # Results area
        self.ics_output = QTextEdit()
        self.ics_output.setReadOnly(True)
        layout.addWidget(self.ics_output)

    def fetch_schedule(self):
        year_study = self.year_combo.currentText()
        term_id = self.term_combo.currentText()
        professor_id = self.teacher_combo.currentData()
        week = int(self.week_combo.currentText())
        
        self.fetch_button.setEnabled(False)
        self.crawler_output.clear()
        self.crawler_progress.setValue(0)
        
        self.crawler_worker = CrawlerWorker(year_study, term_id, professor_id, week)
        self.crawler_worker.finished.connect(self.handle_crawler_result)
        self.crawler_worker.error.connect(self.handle_crawler_error)
        self.crawler_worker.progress.connect(self.update_crawler_progress)
        self.crawler_worker.start()

    def handle_crawler_result(self, schedule):
        try:
            # Generate filename with current date
            current_date = datetime.now().strftime("%Y%m%d")
            filename = f'schedule_{current_date}.json'
            
            # Save to JSON file
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(schedule, f, ensure_ascii=False, indent=2)
            
            # Update table with schedule data
            self.update_schedule_table(schedule)
            
            # Display metadata in a readable format
            metadata = schedule.get('metadata', {})
            metadata_text = (
                f"Week: {metadata.get('week_number', 'N/A')}\n"
                f"Period: {metadata.get('start_date', 'N/A')} - {metadata.get('end_date', 'N/A')}\n"
                f"Professor: {metadata.get('professor_name', 'N/A')}"
            )
            self.metadata_label.setText(metadata_text)
            
            # Display full schedule in output area
            self.crawler_output.setText(json.dumps(schedule, ensure_ascii=False, indent=2))
            self.crawler_progress.setValue(100)
            QMessageBox.information(self, "Success", f"Schedule saved to {filename}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save schedule: {str(e)}")
        finally:
            self.fetch_button.setEnabled(True)

    def update_schedule_table(self, schedule):
        schedule_data = schedule.get('schedule', {})
        all_sessions = []
        
        # Collect all sessions from the schedule
        for day, periods in schedule_data.items():
            for period_type in ['morning', 'afternoon', 'evening']:
                for session in periods.get(period_type, []):
                    if session:  # Check if session exists
                        time_str = f"{day} ({session['time_begin']}-{session['time_end']})"
                        all_sessions.append({
                            'time': time_str,
                            'subject': session['subject'],
                            'room': session['room'],
                            'content': (f"Class: {session['class_name']}\n"
                                      f"Code: {session['class_code']}\n"
                                      f"Period: {session['period']}\n"
                                      f"Taught: {session['taught_lessons']}")
                        })

        # Update table
        self.schedule_table.setRowCount(len(all_sessions))
        for row, session in enumerate(all_sessions):
            self.schedule_table.setItem(row, 0, QTableWidgetItem(session['time']))
            self.schedule_table.setItem(row, 1, QTableWidgetItem(session['subject']))
            self.schedule_table.setItem(row, 2, QTableWidgetItem(session['room']))
            self.schedule_table.setItem(row, 3, QTableWidgetItem(session['content']))
            
            # Make cells read-only
            for col in range(4):
                item = self.schedule_table.item(row, col)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        # Adjust row heights to content
        for row in range(self.schedule_table.rowCount()):
            self.schedule_table.resizeRowToContents(row)

    def handle_crawler_error(self, error_msg):
        self.crawler_output.setText(f"Error: {error_msg}")
        self.crawler_progress.setValue(0)
        self.fetch_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)
        self.schedule_table.setRowCount(0)  # Clear table on error

    def update_crawler_progress(self, message):
        self.crawler_output.append(message)
        # Simulate progress
        current = self.crawler_progress.value()
        self.crawler_progress.setValue(min(current + 30, 90))

    def select_schedule_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select Schedule File",
            "",
            "JSON Files (*.json)"
        )
        if file_name:
            self.file_label.setText(file_name)
            self.export_button.setEnabled(True)

    def export_to_ics(self):
        schedule_file = self.file_label.text()
        if schedule_file == "No file selected":
            QMessageBox.warning(self, "Warning", "Please select a schedule file first")
            return
        
        self.export_button.setEnabled(False)
        self.ics_output.clear()
        self.ics_progress.setValue(0)
        
        self.ics_worker = ICSWorker(schedule_file)
        self.ics_worker.finished.connect(self.handle_ics_result)
        self.ics_worker.error.connect(self.handle_ics_error)
        self.ics_worker.progress.connect(self.update_ics_progress)
        self.ics_worker.start()

    def handle_ics_result(self, ics_content):
        try:
            # Generate filename
            output_file = f"teaching_schedule_{datetime.now().strftime('%Y%m%d')}.ics"
            
            # Save ICS file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(ics_content)
            
            # Display in output area
            self.ics_output.setText(ics_content)
            self.ics_progress.setValue(100)
            QMessageBox.information(self, "Success", f"Calendar exported to {output_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save ICS file: {str(e)}")
        finally:
            self.export_button.setEnabled(True)

    def handle_ics_error(self, error_msg):
        self.ics_output.setText(f"Error: {error_msg}")
        self.ics_progress.setValue(0)
        self.export_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)

    def update_ics_progress(self, message):
        self.ics_output.append(message)
        # Simulate progress
        current = self.ics_progress.value()
        self.ics_progress.setValue(min(current + 45, 90))

    def filter_teachers(self, search_text):
        self.teacher_combo.clear()
        search_text = search_text.lower()
        
        for teacher in self.all_teachers:
            # Search in both full name and ID
            if (search_text in teacher['full_name'].lower() or 
                search_text in teacher['id'].lower()):
                self.teacher_combo.addItem(
                    f"{teacher['full_name']} ({teacher['id']})",
                    teacher['id']
                )
        
        # If we have items after filtering, select the first one
        if self.teacher_combo.count() > 0:
            self.teacher_combo.setCurrentIndex(0)

def main():
    # Create config directory if it doesn't exist
    Path('config').mkdir(exist_ok=True)
    
    # Ensure config files exist with default values if missing
    config_files = {
        'year_studies.json': [{'value': '2023-2024', 'label': '2023-2024'}],
        'terms.json': [{'value': 'HK01', 'label': 'Học kỳ 1'}],
        'teachers.json': [{'id': '011.031.00125', 'full_name': 'Default Teacher'}],
        'weeks.json': [{'value': 1, 'label': 1}]
    }
    
    for filename, default_data in config_files.items():
        config_file = Path('config') / filename
        if not config_file.exists():
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_data, f, ensure_ascii=False, indent=2)
    
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error starting application: {str(e)}")
        raise

if __name__ == "__main__":
    main()
