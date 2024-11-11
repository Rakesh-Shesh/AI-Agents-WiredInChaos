
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import win32com.client as win32

# List of columns to monitor
columns_to_monitor = ["ColumnA", "ColumnB"]  # Replace with actual column names

# Email setup
def send_email(subject, message_body):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = message_body
    mail.To = "your_email@example.com"  # Replace with the recipient's email
    mail.Send()

# Handler to monitor the file for updates
class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self, file_path):
        self.file_path = file_path
        self.last_data = pd.read_excel(file_path)[columns_to_monitor]

    def on_modified(self, event):
        if event.src_path == self.file_path:
            new_data = pd.read_excel(self.file_path)[columns_to_monitor]
            changes = new_data.compare(self.last_data)
            
            if not changes.empty:
                change_message = f"Changes detected in columns {columns_to_monitor}:\n{changes}"
                send_email("Excel Update Alert", change_message)
                self.last_data = new_data

# Set up the observer
def monitor_excel(file_path):
    event_handler = ExcelChangeHandler(file_path)
    observer = Observer()
    observer.schedule(event_handler, path=file_path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

# Start monitoring
excel_file_path = "path_to_your_excel_file.xlsx"  # Replace with your file path
monitor_excel(excel_file_path)
