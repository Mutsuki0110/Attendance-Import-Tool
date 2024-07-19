# Attendance Import Tool

This project is to help facilitate the process of importing missing attendance records of workers.

## Prerequisites

The project is built with `Python` of version `3.9.6`.  
Other `Python` version should also work, but it is not being tested.

## Dependencies

Make sure to install packages listed below before using.

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [shutil](https://docs.python.org/3/library/shutil.html)

## Preparation

The following files should all be in Microsoft Excel spreadsheet format (.xlsx).

1. Manually recorded attendance
2. Exported attendance records from LMS

### Filename Formatting Requirements

You need to manually rename all the files to the specified format, otherwise the code will not work.

1. For weekly attendance record spreadsheet, filename should have a format of `weekly-attendance-{mm}-{dd}.xlsx`.  
For instance, `weekly-attendance-07-18.xlsx` is a valid filename; `weekly_attendance_07_18.xlsx` is not.

2. For attendance records pulled from LMS, filename should have a format of `lms-record-{mm}-{dd}.xlsx`.  
For instance, `lms-record-07-18.xlsx` is a valid filename; `lms_record_07_18.xlsx` is not.

## Usage

1. Put all required files into the `source_files` folder.

## Output Files

By running the code, a combined attendance records spreadsheet will be generated.

1. Inside of the spreadsheet, [PivotTable](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)s will be generated for each individual day.  
All PivotTable spreadsheets will be put in the `pivot_tables` folder.  
The PivotTable contains a column of `Employee ID` and a column of an employee's `daily working hours`.  

- In the main sheet, each day's manual record and record from LMS will be generated using `XLOOKUP` function.
- Records with discrepencies will be highlighted, thus being put into the record-importing file.

The final output file should be a complete Microsoft Excel spreadsheet that contains all the workers' attendance records with specific formats.

- Completing the attendance records importing by uploading the file into LMS.

## Disclaimer

TBC
