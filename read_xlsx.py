'''
This script opens a basic .xlsx file, reads its contents, and displays them to screen
'''
import shutil  # Various utilities including for copying files
import os  # Handles things like deleting files
from datetime import datetime  # For timestamping files
from bs4 import BeautifulSoup  # For picking data out of an XML file - also requires lxml


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

    col_tags = worksheet_soup.find_all('col')
    print(f'There are {len(col_tags)} columns in this sheet:')
    print(col_tags)

    row_tags = worksheet_soup.find_all('row')
    print(f'There are {len(row_tags)} rows in this sheet:')
    print(row_tags)


if __name__ == '__main__':
    file_to_read = r'test_data/test_sheet_01.xlsx'
    temp_archive_path = unzip_xlsx_file(file_to_read)
    sheets_in_workbook = get_sheets_info(temp_archive_path)

    # Extract name and sheetId values
    print(f'There are {len(sheets_in_workbook)} sheets in this workbook')
    for sheet in sheets_in_workbook:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')
        print(f'Sheet Name: {name}, Sheet ID: {sheet_id}')
        read_sheet_contents(name, temp_archive_path)
