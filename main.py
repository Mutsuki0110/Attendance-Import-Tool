import os
import pandas as pd
import shutil
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import datetime, timedelta

# xlookup function
def xlookup(lookup_value, lookup_array, return_array, if_not_found:str = ''):
	match_value = return_array.loc[lookup_array == lookup_value]

	if match_value.empty:
		return '' if if_not_found == '' else if_not_found
	else:
		return match_value.tolist()[0]

# Prompt the user for the initial date
date = input("Enter the date with format 'mm-dd' (e.g., 07-18): ")
initial_file_name = "lms-record-" + date + ".xlsx"

# Extract the date from the file name
base_name = "lms-record-"
date_str = date
initial_date = datetime.strptime(date_str, "%m-%d")

# Define the folder paths
source_files_folder = "source_files"
extracted_files_folder = "extracted_files"

# Create the extracted_files folder if it doesn't exist
os.makedirs(extracted_files_folder, exist_ok=True)

# Loop through all files in the source_files folder that start with "lms-record-"
for file_name in os.listdir(source_files_folder):
	if file_name.startswith("lms-record-") and file_name.endswith(".xlsx"):
		file_path = os.path.join(source_files_folder, file_name)
		
		# Load the Excel file into a DataFrame
		df = pd.read_excel(file_path)
		
		# Extract the specified columns
		extracted_df = df[['employeeNo', 'firstCheckIn', 'lastCheckOut']]
		
		# Save the extracted data to a new Excel file
		extracted_file_name = f"extracted-{file_name}"
		extracted_file_path = os.path.join(extracted_files_folder, extracted_file_name)
		extracted_df.to_excel(extracted_file_path, index=False)
		
		print(f"Processed file: {file_name}, saved extracted data to: {extracted_file_name}")

# Define the folder paths
# pivot_tables_folder = "pivot_tables"
combined_records_folder = "combined_records"
weekly_attendance_file = f"source_files/weekly-attendance-{date}.xlsx"

# Create the combined_records folder if it doesn't exist
os.makedirs(combined_records_folder, exist_ok=True)

# Initialize the Excel writer
merged_file_path = os.path.join(combined_records_folder, "merged-weekly-attendance.xlsx")
with pd.ExcelWriter(merged_file_path) as writer:
	# Read the original weekly attendance file and write it to the merged file
	weekly_attendance_df = pd.read_excel(weekly_attendance_file)
	weekly_attendance_df.to_excel(writer, sheet_name='main', index=False)
	# Read all extracted files and write them to separate sheets
	for file_name in os.listdir(extracted_files_folder):
		if file_name.endswith(".xlsx"):
			file_path = os.path.join(extracted_files_folder, file_name)
			df = pd.read_excel(file_path)
			# Extract the date from the file name
			date_str = file_name.replace("extracted-lms-record-", "").replace(".xlsx", "")
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

# Define the folder paths
combined_records_folder = "combined_records"
merged_file_path = os.path.join(combined_records_folder, "merged-weekly-attendance.xlsx")

# Load the workbook
wb = load_workbook(merged_file_path)
main_df = pd.read_excel(merged_file_path, sheet_name='main')

# Iterate through all the daily record worksheets
for sheet_name in wb.sheetnames:
	if sheet_name != 'main':
		daily_df = pd.read_excel(merged_file_path, sheet_name=sheet_name)
		
		# Perform the XLOOKUP-like operation and add the new column to the main DataFrame
		main_df[f'lms-start-{sheet_name}'] = main_df['employeeNo'].apply(xlookup, args=(daily_df['employeeNo'], daily_df['firstCheckIn']))
		main_df[f'lms-end-{sheet_name}'] = main_df['employeeNo'].apply(xlookup, args=(daily_df['employeeNo'], daily_df['lastCheckOut']))

xlookup_file_path = os.path.join(combined_records_folder, "xlookup-weekly-attendance.xlsx")

# Save the updated main DataFrame back to the merged file
with pd.ExcelWriter(xlookup_file_path) as writer:
	main_df.to_excel(writer, sheet_name='main', index=False)
	for sheet_name in wb.sheetnames:
		if sheet_name != 'main':
			daily_df = pd.read_excel(merged_file_path, sheet_name=sheet_name, engine="openpyxl")
			daily_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"XLOOKUP-ed file saved to: {xlookup_file_path}")

# Define the import template file path
import_template_file = 'final_output_files/importTemplate.xlsx'

# Define the destination directory
import_dest_dir = 'final_output_files'

# Construct the new filename
new_import_filename = f'importTemplate-{date}.xlsx'

# Construct the full destination path
import_dest_path = os.path.join(import_dest_dir, new_import_filename)

# Copy the file to the new destination with the new name
shutil.copy(import_template_file, import_dest_path)

print(f'File copied to {import_dest_path}')

# File paths
xlookup_file_path = 'combined_records/xlookup-weekly-attendance.xlsx'
output_dir = 'final_output_files'

# Load the xlookup weekly attendance workbook
wb = load_workbook(xlookup_file_path)
main_df = pd.read_excel(xlookup_file_path, sheet_name='main')

# Create a list to hold the extracted data
extracted_data = []

# Iterate through all the daily record worksheets
for sheet_name in wb.sheetnames:
	if sheet_name != 'main':
		start_col = f'start-{sheet_name}'
		end_col = f'end-{sheet_name}'
		lms_start_col = f'lms-start-{sheet_name}'
		lms_end_col = f'lms-end-{sheet_name}'

		for _, row in main_df.iterrows():
			if pd.notna(row[lms_start_col]) or pd.notna(row[lms_end_col]):
				continue

			if (pd.notna(row['employeeNo'])) and (pd.isna(row[lms_start_col]) and pd.isna(row[lms_end_col])) and (pd.notna(row[start_col]) and pd.notna(row[end_col])):
				employee_no = row['employeeNo']
				start_time = row[start_col] if pd.isna(row[lms_start_col]) else ''
				end_time = row[end_col] if pd.isna(row[lms_end_col]) else ''
				current_year = datetime.now().year
				date_str = f"{current_year}-{sheet_name}"
				formatted_date = datetime.strptime(date_str, "%Y-%m-%d").strftime("%-m/%d/%Y")
				extracted_data.append([employee_no, "操作员", formatted_date, start_time, end_time, ""])

# Load the import template workbook
import_template_wb = load_workbook(import_template_file)
import_template_ws = import_template_wb.active

# Append the extracted data to the import template
for row in extracted_data:
	import_template_ws.append(row)

# Save the updated import template
new_import_filename = f'importTemplate-{date}.xlsx'
import_template_wb.save(os.path.join(output_dir, new_import_filename))

print(f"Imported records saved to {os.path.join(output_dir, new_import_filename)}")
