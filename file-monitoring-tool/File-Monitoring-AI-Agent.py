"""
This script turns the Excel file monitoring into an AI agent that:
1. Monitors changes in specific columns of an Excel file.
2. Analyzes the data to detect anomalies or trends using a simple machine learning model.
3. Makes intelligent decisions about whether changes are important enough to notify via email.
4. Sends notifications with detailed insights if significant changes are detected.
"""

import pandas as pd  # For reading and manipulating Excel data
from watchdog.observers import Observer  # For monitoring file changes
from watchdog.events import FileSystemEventHandler  # Event handler for file modifications
import time  # For controlling execution flow
import win32com.client as win32  # For sending emails via Outlook
from sklearn.ensemble import IsolationForest  # Anomaly detection model from scikit-learn
import numpy as np  # For numerical operations

# List of columns to monitor
columns_to_monitor = ["ColumnA", "ColumnB"]  # Replace with actual column names


# Email setup function
def send_email(subject, message_body):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = message_body
    mail.To = "your_email@example.com"  # Replace with the recipient's email address
    mail.Send()


# Function to initialize and train a machine learning model for anomaly detection
def train_anomaly_detector(data):
    # Convert data to numerical format if necessary
    numeric_data = data.select_dtypes(include=[np.number])

    # Train an Isolation Forest model (unsupervised anomaly detection)
    model = IsolationForest(contamination=0.05)  # Contamination is the expected percentage of anomalies
    model.fit(numeric_data)

    return model


# Handler class to monitor Excel changes
class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self, file_path):
        self.file_path = file_path
        self.last_data = pd.read_excel(file_path)[columns_to_monitor]

        # Train an initial anomaly detection model on the data
        self.model = train_anomaly_detector(self.last_data)

    def on_modified(self, event):
        if event.src_path == self.file_path:
            # Read the updated data
            new_data = pd.read_excel(self.file_path)[columns_to_monitor]

            # Detect anomalies in the new data
            numeric_data = new_data.select_dtypes(include=[np.number])
            anomalies = self.model.predict(numeric_data)

            # If any anomalies are detected, notify the user
            if np.any(anomalies == -1):  # -1 indicates an anomaly
                change_message = f"Anomalies detected in columns {columns_to_monitor}:\n{new_data[anomalies == -1]}"
                send_email("Excel Update with Anomalies", change_message)

            # Optionally, retrain the model with the new data
            self.model = train_anomaly_detector(new_data)


# Function to start monitoring the Excel file
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


# Start monitoring the Excel file
excel_file_path = "path_to_your_excel_file.xlsx"  # Replace with the actual file path
monitor_excel(excel_file_path)
