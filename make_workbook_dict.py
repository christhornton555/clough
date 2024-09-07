'''
This script takes all the pieces of information from all of the different XML files and assembles them into a dict of 2-d lists
'''

from datetime import datetime, timedelta
import shutil
from get_raw_data_from_xlsx_sheets import unzip_xlsx_file, get_sheets_info, read_sheet_contents
from get_strings_using_string_references import get_string
from convert_sheet_dimensions_to_list import dimensions_to_list, split_cell_reference_into_rows_and_columns
from get_styles_using_style_references import get_style
from get_num_formats_using_num_format_references import make_num_formats_dict


# Convert Excel column codes to integers
def excel_column_to_number(column):
    number = 0
    for char in column:
        number = number * 26 + (ord(char.upper()) - ord('A') + 1)
    return number

# Convert integers to Excel column codes
def number_to_excel_column(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


# Pull the data, styling info, etc from the various XML files it's stored in, and create a dict of tables for the whole workbook
def make_workbook_dict(input_file):
    standard_num_formats_dict = make_num_formats_dict()
    temp_archive_path = unzip_xlsx_file(input_file)

    sheets_in_workbook = get_sheets_info(temp_archive_path)

    workbook = {}

    # Extract the raw data from each sheet
    for sheet in sheets_in_workbook:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')

        workbook[name] = []

        sheet_dimensions, sheet_contents, sheet_column_widths, sheet_row_heights = read_sheet_contents(name, temp_archive_path)

        # print(sheet_column_widths)
        # print(sheet_row_heights)

        number_of_columns = excel_column_to_number(
            dimensions_to_list(sheet_dimensions)[1][0])
        - excel_column_to_number(dimensions_to_list(sheet_dimensions)[0][0]) + 2  # +1 for 1st col & +1 for row/col reference
        number_of_rows = int(
            dimensions_to_list(sheet_dimensions)[1][1])
        - int(dimensions_to_list(sheet_dimensions)[0][1]) + 2  # +1 for 1st col & +1 for row/col reference

        # Set up an array for the sheet, and label the data columns
        workbook[name].append([])
        for c in range(number_of_columns + 1):
            if c == 0:
                workbook[name][0].append('')  # First column is for row numbers, not data. Also avoids zeroth offset problems
            else:
                workbook[name][0].append(number_to_excel_column(c))

        # This is where the magic happens - use the references in each cell to strings & styles to populate the sheet array
        last_row_filled = 0
        for cell in sheet_contents:
            current_col = split_cell_reference_into_rows_and_columns(cell)[0]
            current_row = split_cell_reference_into_rows_and_columns(cell)[1]

            # Create and number a new row when data read in from Excel starts a new row
            if int(current_row) > last_row_filled:
                last_row_filled = int(current_row)
                workbook[name].append([int(current_row)])

            # Find relevant values for cell contents and attributes - TODO - make this into a function?
            current_cell_raw_value = sheet_contents[cell]['raw_value']

            current_cell_style_ref = sheet_contents[cell]['style_num']
            # Convert style ref to actual style
            current_cell_style_dict = get_style(int(current_cell_style_ref), temp_archive_path)
            # print(current_cell_style_dict)
            current_cell_num_style = standard_num_formats_dict[current_cell_style_dict['numFmtId']][1]
            # print(current_cell_num_style)

            # Check if cell has a type specified (e.g. 's' for string)
            if 'type' in sheet_contents[cell]:
                current_cell_type_ref = sheet_contents[cell]['type']
            else:
                current_cell_type_ref = ''
                current_cell_type = ''
            
            current_cell_display_value = current_cell_raw_value

            # If the type is a string, fetch the referenced value
            if current_cell_type_ref == 's':
                current_cell_display_value = get_string(int(sheet_contents[cell]['raw_value']), temp_archive_path)

            # print([cell, current_cell_display_value, current_cell_style_dict])
            # If style is not General (i.e. default, no styling), apply that style (unless cell has no data value to apply it to)
            if current_cell_num_style != 'General' and current_cell_display_value != '':
                if standard_num_formats_dict[current_cell_style_dict['numFmtId']][0] == 'd-mmm-yy':
                    # Excel processes dates idiosynchratically, so we can't just apply the number format without fixing some stuff
                    # TODO - gonna need to apply this to all the other date formats I reckon, ugh. I might have half cracked it below, but need to check
                    # Excel's date system starts on 1900-01-01
                    # Subtract 1 because Excel incorrectly treats 1900 as a leap year
                    base_date = datetime(1899, 12, 30)

                    excel_date = int(current_cell_raw_value)  # Date should be in Excel serial date format, e.g. 45123 = 16/07/2023

                    # Convert the serial date to a datetime object
                    converted_date = base_date + timedelta(days=excel_date)

                    # Apply the desired format
                    current_cell_display_value = converted_date.strftime(current_cell_num_style)

            current_cell_display = f'{current_cell_display_value}'
            workbook[name][int(current_row)].append([current_cell_display, current_cell_style_dict])

    shutil.rmtree(temp_archive_path)  # Tidy up by deleting the temp unzipped archive

    return workbook

if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'
    workbook_to_output = make_workbook_dict(file_to_read)
    print(workbook_to_output)

    print('   --- END ---')
