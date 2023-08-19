#!/bin/env python3

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet as pyxl_Worksheet
import openpyxl.cell.cell
import pandas as pd
import pandas.core.series
import difflib
import sys


def _get_day(curr_day: str, cache: dict = None) -> str:
    """Return the first chronological day present in the given
      string as a lowercase string. Return '' if none is found"""
    if curr_day is None:
        return ''

    if cache is None:
        cache = {}

    if curr_day in cache:
        return cache[curr_day]

    days = ['monday', 'tuesday', 'wednesday', 'thursday',
            'friday', 'saturday', 'sunday']

    curr_day = curr_day.lower()

    for day in days:
        if day in curr_day:
            cache[curr_day] = day.capitalize()
            return cache[curr_day]

    return ''


def parse_timetable(sheet: pyxl_Worksheet) -> pd.DataFrame:
    """Extract course titles, sections, and lecture details from the main
      timetable sheet. Return the extracted data as a pandas dataframe."""
    total_columns = sheet.max_column

    # Starting coordinates of the actual subjects' schedule
    STARTING_ROW, STARTING_COL = 5, 2

    cell_coordinates = {}
    for cell in sheet.merged_cells.ranges:
        cell_coordinates[(cell.min_row, cell.min_col)] = cell.size['columns']

    courses = []
    total_courses = 0
    course_cache = {}

    day = ''
    for row in sheet.iter_rows(min_row=STARTING_ROW):
        if row[0].value is not None:
            day = _get_day(row[0].value)

        room = row[1].value
        if room is None:
            continue
        room = room.strip()

        col_no = STARTING_COL
        while col_no < total_columns:
            cell = row[col_no]
            if cell.value is None:
                col_no += cell_coordinates.get((cell.row, cell.column), 1)
                continue

            if '(' not in cell.value:
                col_no += cell_coordinates.get((cell.row, cell.column), 1)
                continue

            course_details = cell.value.split('(')

            # Ignore anything in brackets within the course name
            # Such things like are of no use to us
            title = course_details[0].split('(')[0].strip()
            # Replace & with and for easier matching later on
            title = title.replace('&', 'and', 1)

            section_list = course_details[1].strip().rstrip(')').split(',')

            start_time = [(8 + (col_no - STARTING_COL) // 6),
                          ((col_no - STARTING_COL) % 6) * 10]

            cell_length = cell_coordinates.get((cell.row, cell.column), 1)
            col_no += cell_length

            length_of_class = [(cell_length // 6),
                               (cell_length % 6) * 10]
            end_time = [start_time[i] + length_of_class[i] for i in range(2)]
            if end_time[1] > 60:
                end_time[0] += 1
                end_time[1] -= 60

            current_lecture = {
                'room': room,
                'day': day,
                'start_time': f'{start_time[0]}:{start_time[1]}',
                'end_time': f'{end_time[0]}:{end_time[1]}',
            }

            for section in section_list:
                section = section.strip()
                if (title, section) in course_cache:
                    courses[course_cache[(title, section)]]['lectures'] \
                        .append(current_lecture)
                else:
                    course_cache[(title, section)] = total_courses
                    total_courses += 1

                    courses.append({
                        'title': title,
                        'section': section,
                        'lectures': [current_lecture],
                    })

    return pd.DataFrame(courses, columns=courses[0].keys())


def _get_dept_from_course_code(course_code: str) -> str:
    """Return the parent department corresponding to the course code.
      Return an empty string ('') for unknown course codes."""
    departments = {'NS': 'NS', ('MT', 'SS', 'SL'): 'HSS',
                   ('CS', 'SE'): 'CS', 'DS': 'DS'}

    if course_code is None:
        return ''
    return departments.get(course_code[:2], '')


def get_course_details(workbook: openpyxl.Workbook,
                       sheets_list: list[pyxl_Worksheet]) -> pd.DataFrame:
    """Extract the course details from all sheets in 'sheets_list' within the
      open workbook. Return a pandas dataframe containing the extracted
      data."""
    course_details = []

    for sheet_name in sheets_list:
        sheet = workbook[sheet_name]
        columns_in_sheet = []

        for cell in sheet[2]:
            if cell.value is None:
                break
            columns_in_sheet.append(cell.value.lower().strip())

        col_num = {}

        # Do we even have the columns ?
        for index, col_name in enumerate(columns_in_sheet):
            if 'code' not in col_num and 'code' in col_name:
                col_num['code'] = index
            elif 'title' not in col_num and 'title' in col_name:
                col_num['title'] = index
            elif 'section' not in col_num and 'section' in col_name:
                col_num['section'] = index
            elif 'instructor' not in col_num and ('teacher' in col_name or
                                                  'instructor' in col_name):
                col_num['instructor'] = index
            elif 'credit_hours' not in col_num and 'credit hour' in col_name:
                col_num['credit_hours'] = index
            elif 'offered_for' not in col_num and 'offered' in col_name:
                col_num['offered_for'] = index
            elif 'category' not in col_num and 'category' in col_name:
                col_num['category'] = index

        # Do we have our main identifiers ?
        if ('code' not in col_num or 'section' not in col_num or
                'title' not in col_num):
            continue

        # First row to start parsing at
        starting_row = 3

        for row in sheet.iter_rows(min_row=starting_row, values_only=True):
            code, title = row[col_num['code']], row[col_num['title']]
            if code is None or title is None:
                continue

            section = row[col_num['section']]
            if section is not None:
                section.strip()

            course = {
                # Replace & with and for easier matching later on,
                'title': title.strip().replace('&', 'and', 1),
                'code': code.strip(),
                'section': section,
            }

            if 'instructor' in col_num:
                instructor = row[col_num['instructor']]
                if instructor is not None:
                    # Ignore VF/CC if mentioned
                    instructor = instructor.split('(')[0].strip()
                course['instructor'] = instructor

            if 'credit_hours' in col_num:
                credit_hours = row[col_num['credit_hours']]
                if type(credit_hours) is not int:
                    credit_hours = None
                course['credit_hours'] = credit_hours

            if 'offered_for' in col_num:
                offered_for = row[col_num['offered_for']]
                if offered_for is not None and type(offered_for) is str:
                    if '(' in offered_for:
                        offered_for = offered_for.split('(')
                        program = offered_for[0].strip()
                        target_dept = offered_for[1].strip().rstrip(')')
                    else:
                        program = offered_for[:2]
                        target_dept = offered_for[2:].strip()
                else:
                    program, target_dept = None, None

                course['program'] = program
                course['target_department'] = target_dept

            if 'category' in col_num:
                category = row[col_num['category']]
                if category is not None and '(' in category:
                    category = category.split('(')
                    parent_dept = category[0].strip()
                    course_type = category[1].strip().lstrip('(').rstrip(')')
                else:
                    parent_dept = _get_dept_from_course_code(code)
                    if parent_dept == '':
                        parent_dept = target_dept
                    course_type = category

                course['parent_department'] = parent_dept
                course['type'] = course_type

            course_details.append(course)

    return pd.DataFrame(course_details, columns=course_details[0].keys())


def _get_corresponding_title(row: pandas.core.series.Series,
                             details_df: pd.DataFrame) -> str:
    """Return the closest matching title for the current row from the
      details DataFrame."""

    title = row['title']
    section = row['section']

    details_exact_match = details_df[
        details_df['section'].apply(lambda sec: sec == row['section'])
    ]
    res = difflib.get_close_matches(title, details_exact_match['title'])

    # Any matches already ?
    if len(res) > 0:
        return res[0]

    # Do we have a different section for the same semester?
    details_semi_match = details_df[
        details_df['section'].apply(
            lambda sec: section[:5] in sec if type(sec) is str else False)
    ]

    res = difflib.get_close_matches(title, details_semi_match['title'])

    # Any luck this time ?
    if len(res) > 0:
        return res[0]

    res = difflib.get_close_matches(title, details_df['title'])

    return res[0] if len(res) > 0 else None


def main():
    # Filename of Excel file
    try:
        filename = sys.argv[1]
    except IndexError:
        filename = 'timetable.xlsx'

    output_csv_filename = f"{filename.split('/')[-1].strip('.xlsx')}.csv"

    # Load the Excel file
    print(f'Attempting to open {filename}')
    try:
        workbook = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        sys.stderr.write(f'Error : Unable to open file: {filename}.\n')
        # _print_example_usage()
        sys.exit(2)

    print(f'Successfully opened {filename}')

    list_of_sheets = workbook.sheetnames
    timetable_sheet = workbook.active.title

    for sheet in list_of_sheets:
        sheet_lower = sheet.lower()
        if 'tt' in sheet_lower or 'timetable' in sheet_lower:
            if timetable_sheet != sheet:
                timetable_sheet = sheet
            break

    list_of_sheets.remove(timetable_sheet)

    print('\nExtracting course details...')
    course_details = get_course_details(workbook, list_of_sheets)
    print('Done.')

    print('\nExtracting class details...')
    course_timetable = parse_timetable(workbook[timetable_sheet])
    print('Done.')

    # Update timetable's course titles to match those in course details
    print('\nMerging course and class details...')
    course_timetable['title'] = course_timetable.apply(
       lambda row: _get_corresponding_title(row, course_details),
       axis=1
    )
    course_data = course_details.merge(
       course_timetable, on=['section', 'title'], how='inner'
    )
    print('Done.')

    print(f'\nExporting to {output_csv_filename}')
    with open(output_csv_filename, 'w') as file:
        course_data.to_csv(file)
    print('Done')

if __name__ == "__main__":
    main()
