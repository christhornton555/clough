'''
This script opens a basic .xlsx file, reads its contents, and displays them to screen
'''
import shutil  # Various utilities including for copying files
import os  # Handles things like deleting files
from datetime import datetime  # For timestamping files
from bs4 import BeautifulSoup  # For picking data out of an XML file


# Return the current time as a string for timestamping filenames
# TODO - I might move this to its own module - I anticipate using it a lot
def get_current_time_string():
    now = datetime.now()
    formatted_time = now.strftime('%Y%m%d_%H%M%S')
    return formatted_time


# Copy any xlsx file to a zip, and then extract
def unzip_xlsx_file(filename):
    # Check the output dir exists, and create it if it doesn't
    current_dir = 'test_data'
    output_dir = f'{current_dir}/output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if filename[-5:] == '.xlsx':
        current_time = get_current_time_string()
        converted_file = f'{current_time}_output.zip'
        zip_filepath = f'{current_dir}/{converted_file}'
        shutil.copy(filename, zip_filepath)  # Create the zip file

        archive_path = f'{output_dir}/{converted_file[:-4]}'
        shutil.unpack_archive(zip_filepath, archive_path)  # Unzips to dir

        os.remove(zip_filepath)  # Get rid of the temp zip file

        print(f'File unzipped to path {archive_path}')

    else:
        print(f'Specified file, {filename}, was not a .xsls file.')

    return archive_path


# Read sheets info from XML
def get_sheets_info(archive_path):
    workbook_path = f'{archive_path}/xl/workbook.xml'
    with open(workbook_path, 'r') as workbook_file:
        workbook_xml_data = workbook_file.read()

    # Create a BeautifulSoup object
    workbook_soup = BeautifulSoup(workbook_xml_data, features='xml')

    # Find all <sheet> tags (N.B. BS ignores the <s:sheet> formatting in the original xml)
    sheet_tags = workbook_soup.find_all('sheet')

    return sheet_tags


# Read contents of individual sheet
def read_sheet_contents(sheet_name, archive_path):
    sheet_path = f'{archive_path}/xl/worksheets/{sheet_name}.xml'
    with open(sheet_path, 'r') as sheet_file:
        sheet_xml_data = sheet_file.read()

    # Create a BeautifulSoup object
    worksheet_soup = BeautifulSoup(sheet_xml_data, features='xml')
    # cell_styles = workbook_soup.find_all('style', attrs={style:family': 'table-cell})

    dimensions_tag = worksheet_soup.find('dimension')
    dimensions = dimensions_tag.get('ref')

    sheetFormatPr_tag = worksheet_soup.find('sheetFormatPr')
    default_row_height = sheetFormatPr_tag.get('defaultRowHeight')

    col_tags = worksheet_soup.find_all('col')
    column_widths = [None]
    for column in range(len(col_tags)):
        # print(f'Column {column} width = {col_tags[column].get('width')}')
        column_widths.append(col_tags[column].get('width'))

    # Read all the raw data into a dictionary
    sheet_dict = {}
    # Find the individual rows in the sheet, and iterate through them
    row_tags = worksheet_soup.find_all('row')
    row_heights = [None]
    for row in range(len(row_tags)):
        if row_tags[row].get('ht') is None:
            # print(f'Row {row_tags[row].get('r')} height = {default_row_height}')
            row_heights.append(default_row_height)
        else:
            # print(f'Row {row_tags[row].get('r')} height = {row_tags[row].get('ht')}')
            row_heights.append(row_tags[row].get('ht'))

        # Find the individual cells in each row, and iterate through them
        cell_tags = row_tags[row].find_all('c')
        for cell in range(len(cell_tags)):
            cell_dict = {
                'raw_value': cell_tags[cell].find('v').text,
                'style_num': cell_tags[cell].get('s')
            }
            if not cell_tags[cell].get('t') is None:  # Some cells have an optional 'Type' property, e.g. 's' = string
                cell_dict.update({'type': cell_tags[cell].get('t')})

            sheet_dict.update({f'{cell_tags[cell].get('r')}': cell_dict})  # Each cell ref is now a dictionary entry

    return dimensions, sheet_dict, column_widths, row_heights
    


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
