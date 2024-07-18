import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import datetime, timedelta

# xlookup function
def xlookup(lookup_value, lookup_array, return_array, if_not_found:str = ''):
    match_value = return_array.loc[lookup_array == lookup_value]

    if match_value.empty:
        return f'"{lookup_value}" not found!' if if_not_found == '' else if_not_found
    else:
        return match_value.tolist()[0]

# Prompt the user for the initial date
start_date = input("Enter the start date with format 'mm-dd' (e.g., 07-18): ")
initial_file_name = "lms-record-" + start_date + ".xlsx"

# Extract the date from the file name
base_name = "lms-record-"
date_str = start_date
initial_date = datetime.strptime(date_str, "%m-%d")

# Loop to generate new file names by incrementing the date by 1 for the next 7 days
for i in range(7):
	new_date = initial_date + timedelta(days=i)
	new_date_str = new_date.strftime("%m-%d")
	new_file_name = f"{base_name}{new_date_str}.xlsx"

	# Load the Excel file into a DataFrame
	df = pd.read_excel(f"source_files/{new_file_name}")

	# Convert "allTime" column to numeric values
	df['allTime'] = pd.to_numeric(df['allTime'], errors='coerce')

	# Create a pivot table with "employeeID" as the index and "allTime" as the values
	pivot_table = pd.pivot_table(df, index='employeeNo', values='allTime', aggfunc='sum')

	# Print the pivot table
	print(f"PivotTable for {new_file_name}:")
	print(pivot_table)

	# Optionally, save the pivot table to a new Excel file
	pivot_table.to_excel(f"pivot_tables/pivottable-{new_file_name}")

	print(f"Processed file: {new_file_name}")

# Define the folder paths
pivot_tables_folder = "pivot_tables"
combined_records_folder = "combined_records"
weekly_attendance_file = f"source_files/weekly-attendance-{start_date}.xlsx"

# Create the combined_records folder if it doesn't exist
os.makedirs(combined_records_folder, exist_ok=True)

# Initialize the Excel writer
merged_file_path = os.path.join(combined_records_folder, "merged-weekly-attendance.xlsx")
with pd.ExcelWriter(merged_file_path) as writer:
    # Read the original weekly attendance file and write it to the merged file
    weekly_attendance_df = pd.read_excel(weekly_attendance_file)
    weekly_attendance_df.to_excel(writer, sheet_name='main', index=False)
    # Read all pivot table files and write them to separate sheets
    for file_name in os.listdir(pivot_tables_folder):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(pivot_tables_folder, file_name)
            df = pd.read_excel(file_path)
            # Extract the date from the file name
            date_str = file_name.replace("pivottable-lms-record-", "").replace(".xlsx", "")
            # Write the DataFrame to a sheet named after the date
            df.to_excel(writer, sheet_name=date_str, index=False)

# Load the workbook to apply date formatting
wb = load_workbook(merged_file_path)
ws = wb['main']

# Define the date format style
date_style = NamedStyle(name="short_date", number_format="MM-DD")

# Apply the date format to the first row
for cell in ws[1]:
    if isinstance(cell.value, datetime):
        cell.style = date_style

# Save the workbook with the updated styles
wb.save(merged_file_path)

print(f"Merged file saved to: {merged_file_path}")
