'''
This script calls all the other modules
'''

from datetime import datetime, timedelta
from get_raw_data_from_xlsx_sheets import unzip_xlsx_file, get_sheets_info, read_sheet_contents
from get_strings_using_string_references import get_string
from convert_sheet_dimensions_to_list import dimensions_to_list
from get_styles_using_style_references import get_style
from get_num_formats_using_num_format_references import make_num_formats_dict

if __name__ == '__main__':
    print('   --- START ---')

    standard_num_formats_dict = make_num_formats_dict()

    file_to_read = r'test_data/test_sheet_01.xlsx'
    temp_archive_path = unzip_xlsx_file(file_to_read)

    sheets_in_workbook = get_sheets_info(temp_archive_path)
    # Extract the raw data from each sheet
    for sheet in sheets_in_workbook:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')
        sheet_dimensions, sheet_contents, sheet_column_widths, sheet_row_heights = read_sheet_contents(name, temp_archive_path)

        # print(sheet_column_widths)
        # print(sheet_row_heights)

        print(f'{name} dimensions: {dimensions_to_list(sheet_dimensions)}')

        for cell in sheet_contents:
            if 'type' in sheet_contents[cell]:
                if sheet_contents[cell]['type'] == 's':
                    print(f'Cell {cell}: {get_string(int(sheet_contents[cell]['raw_value']), temp_archive_path)} (style {sheet_contents[cell]['style_num']}, i.e. numFmt {standard_num_formats_dict[get_style(int(sheet_contents[cell]['style_num']), temp_archive_path)][0]})')
            else:
                if standard_num_formats_dict[get_style(int(sheet_contents[cell]['style_num']), temp_archive_path)][0] == 'General':
                    print(f'Cell {cell}: {sheet_contents[cell]['raw_value']} (style {sheet_contents[cell]['style_num']})')

                elif standard_num_formats_dict[get_style(int(sheet_contents[cell]['style_num']), temp_archive_path)][0] == 'd-mmm-yy':
                    # Excel processes dates idiosynchratically, so we can't just apply the number format without fixing some stuff
                    # TODO - gonna need to apply this to all the other date formats I reckon, ugh. I might have half cracked it below, but need to check
                    # Excel's date system starts on 1900-01-01
                    # Subtract 1 because Excel incorrectly treats 1900 as a leap year
                    base_date = datetime(1899, 12, 30)

                    excel_date = int(sheet_contents[cell]['raw_value'])  # Date should be in Excel serial date format, e.g. 45123 = 16/07/2023

                    # Convert the serial date to a datetime object
                    converted_date = base_date + timedelta(days=excel_date)

                    # Apply the desired format
                    formatted_date = converted_date.strftime(standard_num_formats_dict[get_style(int(sheet_contents[cell]['style_num']), temp_archive_path)][1])
                    
                    print(f'Cell {cell}: {formatted_date} (style {sheet_contents[cell]['style_num']})')

                else:
                    print(f'Cell {cell}: {sheet_contents[cell]['raw_value']} (style {sheet_contents[cell]['style_num']})')

    print('   --- END ---')