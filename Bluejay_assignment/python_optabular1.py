import pandas as pd
from datetime import datetime, timedelta
from tabulate import tabulate
import sys

def analyze_excel_file(file_path, output_file_path):
    # Read Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Perform analysis
    seven_consecutive_days = []
    less_than_10_hours_between_shifts = []
    more_than_14_hours_in_single_shift = []

    current_employee = None
    previous_end_time = None

    for _, row in df.iterrows():
        # Check for 7 consecutive days
        if current_employee and (row['Time'] - previous_end_time).days == 1:
            seven_consecutive_days.append([row['Employee Name'], row['Position Status']])

        # Check for less than 10 hours between shifts
        if current_employee and 1 < (row['Time'] - previous_end_time).seconds // 3600 < 10:
            less_than_10_hours_between_shifts.append([row['Employee Name'], row['Position Status']])

        # Check for more than 14 hours in a single shift
        if (row['Time Out'] - row['Time']).seconds // 3600 > 14:
            more_than_14_hours_in_single_shift.append([row['Employee Name'], row['Position Status']])

        current_employee = row['Employee Name']
        previous_end_time = row['Time Out']

    # Redirect stdout to both console and the output file
    with open(output_file_path, "w") as file:
        original_stdout = sys.stdout
        sys.stdout = file

        # Display results in tabular form
        print("\nEmployees who worked for 7 consecutive days:")
        print(tabulate(seven_consecutive_days, headers=['Employee Name', 'Position Status']))

        print("\nEmployees with less than 10 hours between shifts:")
        print(tabulate(less_than_10_hours_between_shifts, headers=['Employee Name', 'Position Status']))

        print("\nEmployees who worked for more than 14 hours in a single shift:")
        print(tabulate(more_than_14_hours_in_single_shift, headers=['Employee Name', 'Position Status']))

        # Reset stdout to the console
        sys.stdout = original_stdout

if __name__ == "__main__":
    file_path = r"C:\Users\PRATEEK\Downloads\Assignment_Timecard.xlsx"
    output_file_path = r"C:\Users\PRATEEK\Downloads\output.txt"
    analyze_excel_file(file_path, output_file_path)
