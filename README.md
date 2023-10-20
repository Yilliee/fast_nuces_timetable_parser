# FAST NUCES Timetable Parser

This script processes a FAST NUCES timetable, extracting course
details and lecture schedules. The extracted data can be then
outputted in CSV and/or PDF format.

## Dependencies

This script requires that you have python installed on your
system. It depends on openpyxl, pandas and reportlab module.

To install the required dependencies:

```bash
pip install openpyxl pandas reportlab
```

## Usage

To get the exact syntax to run script with, use the following command:

```bash
python3 timetable_parser.py --help
```

If `--excel_file` parameter is provided but no filename is provided next
to it, the script will default to looking for a file named `timetable.xlsx`
in the current working directory.

If the `--create_csv` parameter is passed, a copy of the extracted data will
be saved as a CSV file in the current working directory. The CSV filename
will default to `out.csv` unless overriden via the parameter:
`--output_csv <csv_name>`.

If the `--create_pdf` parameter is passed, a copy of the extracted data will
be saved as a pdf file in the current working directory. The PDF filename
will default to `out.pdf` unless overriden via the parameter:
 `--output_pdf <pdf_name>`.

## Example

The Fall 2023 v1.1 timetable is provided as a sample file in the
sample folder. Run the script for it with using the following
command:

```bash
python3 timetable_parser.py --excel_file sample/FSC_Time_Table__List_of_Courses_Fall_2023_v1.1.xlsx --create_pdf --output_pdf "FSC_TimeTable_Fall23_v1.1.pdf"
```

The output file would be saved in the current directory with the
name: `FSC_Timetable_Fall23_1.1.pdf`

## Features

- Extracts course details, lecture schedules and instructor names
  from the FAST NUCES FSC timetable excel file.
- Automatically links data from timetable sheet with that from the
  list of courses sheets.
- Allows to export the processed data as a pdf as well as a csv file.

## Compatibility

Currently tested only on the FSC (FAST SCHOOL OF COMPUTING)
timetable for Lahore campus ( Fall22, Spring23, Fall23 ).
Other timetables should work as long as they follow a similar
format with minor edits.

## Disclaimer

This script is provided as-is and may require adjustments for
compatibility with different timetables or campuses.
Use at your own discretion.

## License

This project is licensed under the [GNU GPL v3.0 License](LICENSE).
