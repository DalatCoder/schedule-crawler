import sys
import json
from datetime import datetime
from pathlib import Path
import os
import subprocess

# Add platform-specific imports for opening file explorer
from sys import platform

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QComboBox, QPushButton, 
    QTextEdit, QTabWidget, QFileDialog, QMessageBox,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from schedule_crawler import ScheduleCrawler
from ics_exporter import ICSExporter

class CrawlerWorker(QThread):
    """Worker thread for crawling schedule data asynchronously"""
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
        """Execute the crawler operation in a separate thread to avoid blocking UI"""
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
    """Worker thread for ICS file generation operations"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, schedule_file):
        super().__init__()
        self.schedule_file = schedule_file

    def run(self):
        """Execute the ICS file generation operation in a separate thread to avoid blocking UI"""
        try:
            self.progress.emit("Creating ICS file...")
            exporter = ICSExporter()
            ics_content = exporter.create_ics_content(self.schedule_file)
            self.progress.emit("ICS content generated successfully!")
            self.finished.emit(ics_content)
        except Exception as e:
            self.error.emit(str(e))

    def create_from_data(self, schedule_data):
        """Create ICS content directly from schedule data without reading from file"""
        try:
            self.progress.emit("Creating ICS file...")
            exporter = ICSExporter()
            ics_content = exporter.create_ics_content_from_data(schedule_data)
            self.progress.emit("ICS content generated successfully!")
            self.finished.emit(ics_content)
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    """Main application window containing all UI components and logic"""
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
        
        # Create Crawler tabs
        teacher_tab = QWidget()
        teacher_layout = QVBoxLayout(teacher_tab)
        tabs.addTab(teacher_tab, "Teacher Schedule")
        
        student_tab = QWidget()
        student_layout = QVBoxLayout(student_tab)
        tabs.addTab(student_tab, "Student Schedule")
        
        # Create ICS Exporter tab
        ics_tab = QWidget()
        ics_layout = QVBoxLayout(ics_tab)
        tabs.addTab(ics_tab, "ICS Exporter")
        
        # Setup tabs
        self.setup_crawler_tab(teacher_layout)  # Rename existing crawler tab setup
        self.setup_student_tab(student_layout)  # Add new student tab setup
        self.setup_ics_tab(ics_layout)
        
        # Load configuration data
        self.load_config_data()
        self.load_student_config_data()

        self.current_schedule = None
        self.current_student_schedule = None
        self.last_ics_file = None

    def load_config_data(self):
        """Load configuration data from JSON files including years, terms, teachers, and weeks"""
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
        """Initialize and setup the crawler tab with all its UI components"""
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
        button_layout = QHBoxLayout()
        self.fetch_button = QPushButton("Fetch Schedule")
        self.fetch_button.clicked.connect(self.fetch_schedule)
        button_layout.addWidget(self.fetch_button)

        self.export_ics_button = QPushButton("Export to ICS")
        self.export_ics_button.clicked.connect(self.export_current_to_ics)
        self.export_ics_button.setEnabled(False)  # Disabled by default
        button_layout.addWidget(self.export_ics_button)
        form_layout.addLayout(button_layout)
        
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

    def setup_student_tab(self, layout):
        """Initialize and setup the student schedule crawler tab"""
        # Create form layout for inputs
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        layout.addWidget(form_widget)
        
        # Today's date
        today_layout = QHBoxLayout()
        today_layout.addWidget(QLabel("Today:"))
        self.student_today_label = QLabel(datetime.now().strftime("%d/%m/%Y"))
        today_layout.addWidget(self.student_today_label)
        form_layout.addLayout(today_layout)
        
        # Schedule metadata section
        metadata_layout = QHBoxLayout()
        metadata_layout.addWidget(QLabel("Schedule Info:"))
        self.student_metadata_label = QLabel("No schedule loaded")
        self.student_metadata_label.setWordWrap(True)
        metadata_layout.addWidget(self.student_metadata_label)
        form_layout.addLayout(metadata_layout)
        
        # Year selection
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel("Academic Year:"))
        self.student_year_combo = QComboBox()
        year_layout.addWidget(self.student_year_combo)
        form_layout.addLayout(year_layout)
        
        # Term selection
        term_layout = QHBoxLayout()
        term_layout.addWidget(QLabel("Term:"))
        self.student_term_combo = QComboBox()
        term_layout.addWidget(self.student_term_combo)
        form_layout.addLayout(term_layout)
        
        # Class selection with filter
        class_layout = QVBoxLayout()
        class_header = QHBoxLayout()
        class_header.addWidget(QLabel("Class:"))
        
        # Add search box for class
        self.class_filter = QLineEdit()
        self.class_filter.setPlaceholderText("Search class...")
        self.class_filter.textChanged.connect(self.filter_classes)
        class_header.addWidget(self.class_filter)
        
        class_layout.addLayout(class_header)
        
        self.class_combo = QComboBox()
        self.class_combo.setMaxVisibleItems(10)
        class_layout.addWidget(self.class_combo)
        form_layout.addLayout(class_layout)
        
        # Week selection
        week_layout = QHBoxLayout()
        week_layout.addWidget(QLabel("Week:"))
        self.student_week_combo = QComboBox()
        week_layout.addWidget(self.student_week_combo)
        form_layout.addLayout(week_layout)
        
        # Fetch button
        button_layout = QHBoxLayout()
        self.student_fetch_button = QPushButton("Fetch Schedule")
        self.student_fetch_button.clicked.connect(self.fetch_student_schedule)
        button_layout.addWidget(self.student_fetch_button)

        self.student_export_ics_button = QPushButton("Export to ICS")
        self.student_export_ics_button.clicked.connect(self.export_current_student_to_ics)
        self.student_export_ics_button.setEnabled(False)
        button_layout.addWidget(self.student_export_ics_button)
        form_layout.addLayout(button_layout)
        
        # Progress bar
        self.student_progress = QProgressBar()
        self.student_progress.setTextVisible(True)
        self.student_progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.student_progress)
        
        # Schedule Table
        self.student_table = QTableWidget()
        self.student_table.setColumnCount(4)
        self.student_table.setHorizontalHeaderLabels(['Time', 'Subject', 'Room', 'Content'])
        self.student_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.student_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.student_table)
        
        # Results area
        self.student_output = QTextEdit()
        self.student_output.setReadOnly(True)
        self.student_output.setMaximumHeight(100)
        layout.addWidget(self.student_output)

    def load_student_config_data(self):
        """Load student-specific configuration data"""
        try:
            # Load years (reuse existing years)
            with open('student_config/year_studies.json', 'r', encoding='utf-8') as f:
                years = json.load(f)
                self.student_year_combo.addItems([year['value'] for year in years])
            
            # Load terms
            with open('student_config/terms.json', 'r', encoding='utf-8') as f:
                terms = json.load(f)
                self.student_term_combo.addItems([term['value'] for term in terms])
            
            # Load classes
            with open('student_config/classes.json', 'r', encoding='utf-8') as f:
                self.all_classes = json.load(f)
                for class_item in self.all_classes:
                    self.class_combo.addItem(class_item['value'])
            
            # Load weeks
            with open('student_config/weeks.json', 'r', encoding='utf-8') as f:
                weeks = json.load(f)
                self.student_week_combo.addItems([str(week['label']) for week in weeks])
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load student configuration: {str(e)}")

    def fetch_schedule(self):
        """Initiate the schedule crawling process with selected parameters"""
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
        """Process and display the crawled schedule data"""
        try:
            # Store the current schedule
            self.current_schedule = schedule
            
            # Enable the export button
            self.export_ics_button.setEnabled(True)
            
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
            self.current_schedule = None
            self.export_ics_button.setEnabled(False)
            QMessageBox.critical(self, "Error", f"Failed to save schedule: {str(e)}")
        finally:
            self.fetch_button.setEnabled(True)

    def update_schedule_table(self, schedule):
        """Update the UI table with the fetched schedule data"""
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
        """Handle and display any errors that occur during crawling"""
        self.current_schedule = None
        self.export_ics_button.setEnabled(False)
        self.crawler_output.setText(f"Error: {error_msg}")
        self.crawler_progress.setValue(0)
        self.fetch_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)
        self.schedule_table.setRowCount(0)  # Clear table on error

    def update_crawler_progress(self, message):
        """Update the progress bar and display crawling status messages"""
        self.crawler_output.append(message)
        # Simulate progress
        current = self.crawler_progress.value()
        self.crawler_progress.setValue(min(current + 30, 90))

    def select_schedule_file(self):
        """Open file dialog for selecting a saved schedule JSON file"""
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
        """Export selected schedule file to ICS format"""
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

    def export_current_to_ics(self):
        """Export currently loaded schedule data to ICS format"""
        if not self.current_schedule:
            QMessageBox.warning(self, "Warning", "No schedule data available")
            return
        
        self.export_ics_button.setEnabled(False)
        self.ics_progress.setValue(0)
        
        # Create a new worker for ICS export
        self.ics_worker = ICSWorker(None)  # No file needed
        self.ics_worker.finished.connect(self.handle_ics_result)
        self.ics_worker.error.connect(self.handle_ics_error)
        self.ics_worker.progress.connect(self.update_ics_progress)
        
        # Use create_from_data instead of run
        self.ics_worker.create_from_data(self.current_schedule)

    def handle_ics_result(self, ics_content):
        """Process and save the generated ICS content"""
        try:
            # Generate filename
            output_file = f"teaching_schedule_{datetime.now().strftime('%Y%m%d')}.ics"
            
            # Save ICS file with full path
            current_dir = os.getcwd()
            full_path = os.path.join(current_dir, output_file)
            
            with open(full_path, 'w', encoding='utf-8') as f:
                f.write(ics_content)
            
            # Store the last ICS file path
            self.last_ics_file = full_path
            
            # Display in output area
            self.ics_output.setText(ics_content)
            self.ics_progress.setValue(100)
            
            # Create result message box with option to open location
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Success")
            msg_box.setText(f"Calendar exported to {output_file}")
            msg_box.setIcon(QMessageBox.Icon.Information)
            
            # Add custom button to open file location
            open_location_button = msg_box.addButton("Open Location", QMessageBox.ButtonRole.ActionRole)
            close_button = msg_box.addButton(QMessageBox.StandardButton.Close)
            
            msg_box.exec()
            
            # Handle button click
            if msg_box.clickedButton() == open_location_button:
                self.open_file_location(full_path)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save ICS file: {str(e)}")
        finally:
            self.export_button.setEnabled(True)
            self.export_ics_button.setEnabled(True)

    def open_file_location(self, file_path):
        """Open the system file explorer at the specified file location"""
        try:
            if platform == "win32":
                # Windows
                subprocess.run(['explorer', '/select,', os.path.normpath(file_path)])
            elif platform == "darwin":
                # macOS
                subprocess.run(['open', '-R', file_path])
            else:
                # Linux
                subprocess.run(['xdg-open', os.path.dirname(file_path)])
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Could not open file location: {str(e)}")

    def handle_ics_error(self, error_msg):
        """Handle and display any errors that occur during ICS export"""
        self.ics_output.setText(f"Error: {error_msg}")
        self.ics_progress.setValue(0)
        self.export_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)

    def update_ics_progress(self, message):
        """Update the progress bar and display ICS export status messages"""
        self.ics_output.append(message)
        # Simulate progress
        current = self.ics_progress.value()
        self.ics_progress.setValue(min(current + 45, 90))

    def filter_teachers(self, search_text):
        """Filter the teachers dropdown list based on search text"""
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

    def filter_classes(self, search_text):
        """Filter the classes dropdown list based on search text"""
        self.class_combo.clear()
        search_text = search_text.lower()
        
        for class_item in self.all_classes:
            if search_text in class_item['value'].lower():
                self.class_combo.addItem(class_item['value'])
        
        if self.class_combo.count() > 0:
            self.class_combo.setCurrentIndex(0)

    def fetch_student_schedule(self):
        """Fetch student schedule with selected parameters"""
        year_study = self.student_year_combo.currentText()
        term_id = self.student_term_combo.currentText()
        class_id = self.class_combo.currentText()
        week = int(self.student_week_combo.currentText())
        
        self.student_fetch_button.setEnabled(False)
        self.student_output.clear()
        self.student_progress.setValue(0)
        
        # Create and start student crawler worker
        from student_schedule_crawler import StudentScheduleCrawler
        crawler = StudentScheduleCrawler()
        crawler.year_study = year_study
        crawler.term_id = term_id
        crawler.class_id = class_id
        
        try:
            schedule = crawler.fetch_schedule(week)
            self.handle_student_crawler_result(schedule)
        except Exception as e:
            self.handle_student_crawler_error(str(e))

    def handle_student_crawler_result(self, schedule):
        """Process and display the crawled student schedule data"""
        try:
            # Store the current schedule
            self.current_student_schedule = schedule
            
            # Enable the export button
            self.student_export_ics_button.setEnabled(True)
            
            # Generate filename with current date
            current_date = datetime.now().strftime("%Y%m%d")
            filename = f'student_schedule_{current_date}.json'
            
            # Save to JSON file
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(schedule, f, ensure_ascii=False, indent=2)
            
            # Update table with schedule data
            self.update_student_schedule_table(schedule)
            
            # Display metadata in a readable format
            metadata = schedule.get('metadata', {})
            metadata_text = (
                f"Week: {metadata.get('week_number', 'N/A')}\n"
                f"Period: {metadata.get('start_date', 'N/A')} - {metadata.get('end_date', 'N/A')}\n"
                f"Class: {metadata.get('class_name', 'N/A')}"
            )
            self.student_metadata_label.setText(metadata_text)
            
            # Display full schedule in output area
            self.student_output.setText(json.dumps(schedule, ensure_ascii=False, indent=2))
            self.student_progress.setValue(100)
            QMessageBox.information(self, "Success", f"Schedule saved to {filename}")
        except Exception as e:
            self.current_student_schedule = None
            self.student_export_ics_button.setEnabled(False)
            QMessageBox.critical(self, "Error", f"Failed to save schedule: {str(e)}")
        finally:
            self.student_fetch_button.setEnabled(True)

    def handle_student_crawler_error(self, error_msg):
        """Handle and display any errors that occur during student schedule crawling"""
        self.current_student_schedule = None
        self.student_export_ics_button.setEnabled(False)
        self.student_output.setText(f"Error: {error_msg}")
        self.student_progress.setValue(0)
        self.student_fetch_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)
        self.student_table.setRowCount(0)

    def update_student_schedule_table(self, schedule):
        """Update the UI table with the fetched student schedule data"""
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
                            'content': (f"Teacher: {session['teacher_name']}\n"
                                      f"Code: {session['class_code']}\n"
                                      f"Period: {session['period']}")
                        })

        # Update table
        self.student_table.setRowCount(len(all_sessions))
        for row, session in enumerate(all_sessions):
            self.student_table.setItem(row, 0, QTableWidgetItem(session['time']))
            self.student_table.setItem(row, 1, QTableWidgetItem(session['subject']))
            self.student_table.setItem(row, 2, QTableWidgetItem(session['room']))
            self.student_table.setItem(row, 3, QTableWidgetItem(session['content']))
            
            # Make cells read-only
            for col in range(4):
                item = self.student_table.item(row, col)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        # Adjust row heights to content
        for row in range(self.student_table.rowCount()):
            self.student_table.resizeRowToContents(row)

    def export_current_student_to_ics(self):
        """Export currently loaded student schedule data to ICS format"""
        if not self.current_student_schedule:
            QMessageBox.warning(self, "Warning", "No student schedule data available")
            return
        
        self.student_export_ics_button.setEnabled(False)
        self.student_progress.setValue(0)
        
        # Create a new worker for ICS export
        from student_ics_exporter import StudentICSExporter
        exporter = StudentICSExporter()  # Create StudentICSExporter instance
        self.student_ics_worker = ICSWorker(None)
        self.student_ics_worker.finished.connect(self.handle_student_ics_result)
        self.student_ics_worker.error.connect(self.handle_student_ics_error)
        self.student_ics_worker.progress.connect(self.update_student_ics_progress)
        
        try:
            # Use StudentICSExporter's methods directly
            ics_content = exporter.create_ics_content_from_data(self.current_student_schedule)
            self.student_ics_worker.finished.emit(ics_content)
        except Exception as e:
            self.student_ics_worker.error.emit(str(e))

    def handle_student_ics_result(self, ics_content):
        """Process and save the generated student ICS content"""
        try:
            # Generate filename
            output_file = f"student_schedule_{datetime.now().strftime('%Y%m%d')}.ics"
            
            # Save ICS file with full path
            current_dir = os.getcwd()
            full_path = os.path.join(current_dir, output_file)
            
            with open(full_path, 'w', encoding='utf-8') as f:
                f.write(ics_content)
            
            # Display in output area
            self.student_output.setText(ics_content)
            self.student_progress.setValue(100)
            
            # Create result message box with option to open location
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Success")
            msg_box.setText(f"Calendar exported to {output_file}")
            msg_box.setIcon(QMessageBox.Icon.Information)
            
            # Add custom button to open file location
            open_location_button = msg_box.addButton("Open Location", QMessageBox.ButtonRole.ActionRole)
            close_button = msg_box.addButton(QMessageBox.StandardButton.Close)
            
            msg_box.exec()
            
            # Handle button click
            if msg_box.clickedButton() == open_location_button:
                self.open_file_location(full_path)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save ICS file: {str(e)}")
        finally:
            self.student_export_ics_button.setEnabled(True)

    def handle_student_ics_error(self, error_msg):
        """Handle and display any errors that occur during student ICS export"""
        self.student_output.setText(f"Error: {error_msg}")
        self.student_progress.setValue(0)
        self.student_export_ics_button.setEnabled(True)
        QMessageBox.critical(self, "Error", error_msg)

    def update_student_ics_progress(self, message):
        """Update the progress bar and display student ICS export status messages"""
        self.student_output.append(message)
        # Simulate progress
        current = self.student_progress.value()
        self.student_progress.setValue(min(current + 45, 90))

    def setup_ics_tab(self, layout):
        """Initialize and setup the ICS export tab"""
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
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export_to_ics)
        layout.addWidget(self.export_button)
        
        # Progress bar
        self.ics_progress = QProgressBar()
        self.ics_progress.setTextVisible(True)
        self.ics_progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.ics_progress)
        
        # ICS output area
        self.ics_output = QTextEdit()
        self.ics_output.setReadOnly(True)
        layout.addWidget(self.ics_output)

def main():
    """Application entry point - sets up configuration and launches the UI"""
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
