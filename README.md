# Attendance Import Tool

This project is to help facilitate the process of importing missing attendance records of workers.

## Preparation

The following files should all be in Microsoft Excel spreadsheet format (.xlsx).

1. Weekly recorded attendance
2. Daily exported attendance records from LMS

All files should be put into the `source_files` folder.

## Output Files

By running the code, a combined attendance records spreadsheet should be generated.

- Inside of the spreadsheet, PivotCharts will be generated for each individual day.
- Inside of the PivotChart, there is a column of Employee ID and a column of an employee's daily working hours.
- In the main sheet, each day's manual record and record from LMS will be generated using XLOOKUP function.
- Records with discrepencies will be highlighted, thus being put into the record-importing file.

The final output file should be a complete Microsoft Excel spreadsheet that contains all the workers' attendance records with specific formats.

- Completing the attendance records importing by uploading the file into LMS.

## Dependencies

1. pandas
