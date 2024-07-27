# Attendance Import Tool

This project is to help facilitate the process of importing missing attendance records of workers.

## Prerequisites

The project is built with `Python` of version `3.9.6`.  
Other `Python` version should also work, but it is not being tested.

## Dependencies

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Streamlit](https://streamlit.io/)

## Preparation

The following files should all be in Microsoft Excel spreadsheet format (.xlsx).

1. Manually recorded attendance
2. Exported attendance records from LMS

### File Name Formatting Requirements

You need to manually rename all the files to the specified format, otherwise the code will not work.

1. For manual attendance record spreadsheet, filename should have a format of `manual-attendance-{mm}-{dd}.xlsx`.  
For instance, `manual-attendance-07-18.xlsx` is a valid filename; `manual_attendance_07_18.xlsx` is not.

2. For attendance records pulled from LMS, filename should have a format of `lms-record-{mm}-{dd}.xlsx`.  
For instance, `lms-record-07-18.xlsx` is a valid filename; `lms_record_07_18.xlsx` is not.

## Usage

1. Create a folder named `source_files` in the root directory.
2. Put all required files into the `source_files` folder.  
For instance, `manual-attendance-07-18.xlsx` and `lms-record-07-18.xlsx`.
3. Run the code.
4. The final output files will be generated in the `final_output_files` directory.
5. Import the final output files into LMS.

## Output Files

By running the code, two additional directories will be created in the root directory: `combined_records` and `extracted_files`.

1. `combined_records` - This directory contains the merged and XLOOKUP-ed attendance records of both manual and LMS records.
2. `extracted_files` - This directory contains the extracted check-in and check-out time from the attendance records provided.

## Disclaimer

This project is not affiliated with **JD Logistics United States Company** and is not under the regulation of **JD Logistics United States Company**. By using this project, you agree that any unwanted results or consequences are not the responsibility of the author of this project.
