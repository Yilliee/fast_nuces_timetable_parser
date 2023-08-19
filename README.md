# FAST NUCES Timetable Parser

This script processes a FAST NUCES timetable, extracting course
details and lecture schedules. The extracted data is then
outputted in CSV format.

## Dependencies

This script requires that you have python installed on your
system. It depends on openpyxl and pandas modules.

To install the required dependencies:

```bash
pip install openpyxl pandas
```

## Usage

Run the script with the following command:

```bash
python3 timetable_parser.py [path_to_excel_file]
```

If no Excel file path is provided, the script will default to
looking for a file named `timetable.xlsx` in the same
directory.

The extracted data will be saved as a CSV file in the current
working directory. The CSV filename will match the input
Excel file's name.

## Example

The Fall 2023 v1.1 timetable is provided as a sample file in the
sample folder. Run the script for it with using the following
command:

```bash
python3 timetable_parser.py sample/FSC_Time_Table__List_of_Courses_Fall_2023_v1.1.xlsx
```

The output file would be saved in the current directory with the
name: `FSC_Time_Table__List_of_Courses_Fall_2023_v1.1.csv`

## Features

- Extracts course details and lecture schedules from the FAST
  NUCES timetable.
- Automatically matches course titles between timetable and
  course details.

## Compatibility

Currently tested only on the FSC (FAST SCHOOL OF COMPUTING)
timetable for Lahore campus ( Fall22, Spring23, Fall23 ).
Other timetables should work as long as they follow a similar
format.

## Disclaimer

This script is provided as-is and may require adjustments for
compatibility with different timetables or campuses.
Use at your own discretion.

## License

This project is licensed under the [GNU GPL v3.0 License](LICENSE).
