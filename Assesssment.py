import pandas as pd
from datetime import timedelta

def analyze_excel_file(file_path):
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)

    # Convert date columns to datetime type
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        # Check for employees who have worked for 7 consecutive days
        consecutive_days = df[(df['Employee Name'] == row['Employee Name']) &
                              (df['Time'] >= row['Time']) &
                              (df['Time'] <= row['Time'] + timedelta(days=6))]
        if len(consecutive_days) >= 7:
            print(f"Employee {row['Employee Name']} has worked for 7 consecutive days at {row['Time']}")

        # Check for employees with less than 10 hours between shifts but greater than 1 hour
        next_shift = df[(df['Employee Name'] == row['Employee Name']) &
                        (df['Time'] > row['Time Out'])]
        time_between_shifts = next_shift['Time'] - row['Time Out']
        if not time_between_shifts.empty and (time_between_shifts.min() < pd.Timedelta('10 hours') and
                                              time_between_shifts.min() > pd.Timedelta('1 hour')):
            print(f"Employee {row['Employee Name']} has less than 10 hours between shifts at {row['Time']}")

        # Check for employees who have worked for more than 14 hours in a single shift
        if (row['Time Out'] - row['Time']) > pd.Timedelta('14 hours'):
            print(f"Employee {row['Employee Name']} has worked for more than 14 hours at {row['Time']}")

if __name__ == "__main__":
    # Assuming the Excel file is named 'employee_schedule.xlsx' and is in the current directory
    file_path = 'Assignment_Timecard.xlsx'
    analyze_excel_file(file_path)
