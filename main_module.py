from get_raw_data_from_xlsx_sheets import unzip_xlsx_file, get_sheets_info, read_sheet_contents


if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'

    temp_archive_path = unzip_xlsx_file(file_to_read)
    sheets_in_workbook = get_sheets_info(temp_archive_path)

    # Extract name and sheetId values
    if len(sheets_in_workbook) == 1:
        print(f'There is {len(sheets_in_workbook)} sheet in this workbook')
    else:
        print(f'There are {len(sheets_in_workbook)} sheets in this workbook')
    
    # Extract the raw data from each sheet
    for sheet in sheets_in_workbook:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')
        print(f'Sheet Name: {name}, Sheet ID: {sheet_id}')
        sheet_dimensions, sheet_contents, sheet_column_widths, sheet_row_heights = read_sheet_contents(name, temp_archive_path)

        # Display the data extracted from the sheet
        print(f'Sheet dimensions = {sheet_dimensions}')
        cell_reference = 'A1'
        print(f'The raw values of cell {cell_reference} are: {sheet_contents[cell_reference]}')
        # N.B. Excel numbers values starting at 1 (or A), so the zeroth width and height are None to reduce the chance of errors
        print(f'Column widths of {name} in order are: {sheet_column_widths}')
        print(f'Row heights of {name} in order are: {sheet_row_heights}')

    print('   --- END ---')