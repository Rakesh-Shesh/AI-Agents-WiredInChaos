"""
This script monitors an Excel file for updates in specific columns (e.g., "ColumnA" and "ColumnB").
It uses the `watchdog` library to monitor the file system for changes to the Excel file.
When a change is detected, it compares the new data with the previously stored data from the specified columns.

If any changes are found in the monitored columns, an email alert is sent using the `win32com.client`
module to send an email via Outlook. The email contains the details of the changes detected in the specified columns.

Steps:
1. Monitor changes in the Excel file.
2. If a change is detected, compare it with the last known data.
3. If any changes in the monitored columns are found, send an email with the details.
4. The program continuously monitors the file until manually interrupted.

Note: The script assumes you are using Outlook to send emails and `pandas` to process the Excel file.
"""

import pandas as pd  # Import pandas for handling Excel files and data manipulation
from watchdog.observers import Observer  # Import Observer from watchdog to monitor file system changes
from watchdog.events import FileSystemEventHandler  # Import FileSystemEventHandler to handle file changes
import time  # Import time to control the execution frequency
import win32com.client as win32  # Import win32com to send emails through Outlook

# List of columns to monitor for changes
columns_to_monitor = ["ColumnA", "ColumnB"]  # Replace these with the actual column names in your Excel file


# Email setup function using Outlook to send an email when changes are detected
def send_email(subject, message_body):
    # Create an Outlook application instance to send the email
    outlook = win32.Dispatch("outlook.application")

    # Create a new mail item
    mail = outlook.CreateItem(0)

    # Set up the subject and body of the email
    mail.Subject = subject  # Set the subject of the email
    mail.Body = message_body  # Set the body of the email, containing the change details

    # Specify the recipient's email address
    mail.To = "your_email@example.com"  # Replace this with the recipient's email address

    # Send the email
    mail.Send()


# Excel change handler class to detect changes in specific columns
class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self, file_path):
        # Initialize the handler with the path of the Excel file to monitor
        self.file_path = file_path

        # Read the Excel file and store the initial data of the columns to monitor
        self.last_data = pd.read_excel(file_path)[columns_to_monitor]

    # Function to handle the file modification event
    def on_modified(self, event):
        # Check if the modified file is the one being monitored
        if event.src_path == self.file_path:
            # Read the updated data from the Excel file
            new_data = pd.read_excel(self.file_path)[columns_to_monitor]

            # Compare the new data with the last known data to detect any changes
            changes = new_data.compare(self.last_data)

            # If there are any changes detected, send an email
            if not changes.empty:
                # Create a message to describe the changes detected in the monitored columns
                change_message = f"Changes detected in columns {columns_to_monitor}:\n{changes}"

                # Send an email notification with the detected changes
                send_email("Excel Update Alert", change_message)

                # Update the stored last data to the new data after sending the email
                self.last_data = new_data


# Function to set up the file monitoring process
def monitor_excel(file_path):
    # Create an instance of the ExcelChangeHandler class, passing the file path to it
    event_handler = ExcelChangeHandler(file_path)

    # Set up the observer to monitor changes in the file system
    observer = Observer()

    # Schedule the observer to watch the specified file for changes, with no recursion into subdirectories
    observer.schedule(event_handler, path=file_path, recursive=False)

    # Start the observer to begin monitoring the file
    observer.start()

    # Keep the program running to monitor for changes until manually interrupted
    try:
        while True:
            # Sleep for 1 second to prevent high CPU usage and keep checking for events
            time.sleep(1)
    except KeyboardInterrupt:
        # If the user presses Ctrl+C, stop the observer
        observer.stop()

    # Join the observer thread to allow it to finish running
    observer.join()


# Start monitoring the Excel file by specifying its path
excel_file_path = "path_to_your_excel_file.xlsx"  # Replace with the actual path to the Excel file
monitor_excel(excel_file_path)  # Call the monitor function to start monitoring the Excel file
