'''
This script calls all the other modules
'''

from get_raw_data_from_xlsx_sheets import unzip_xlsx_file, get_sheets_info, read_sheet_contents
from get_strings_using_string_references import get_string

if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'
    temp_archive_path = unzip_xlsx_file(file_to_read)

    sheets_in_workbook = get_sheets_info(temp_archive_path)
    # Extract the raw data from each sheet
    for sheet in sheets_in_workbook:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')
        sheet_dimensions, sheet_contents, sheet_column_widths, sheet_row_heights = read_sheet_contents(name, temp_archive_path)

        # print(name)
        # print(sheet_dimensions)
        # print(sheet_column_widths)
        # print(sheet_row_heights)

        cell_ref = 'B1'
        if 'type' in sheet_contents[cell_ref]:
            if sheet_contents[cell_ref]['type'] == 's':
                print(get_string(int(sheet_contents[cell_ref]['raw_value']), temp_archive_path))
        else:
            print(sheet_contents[cell_ref])

    print('   --- END ---')