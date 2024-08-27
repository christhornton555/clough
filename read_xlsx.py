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
    if filename[-5:] == '.xlsx':
        current_time = get_current_time_string()
        converted_file = f'test_data/{current_time}_output.zip'
        shutil.copy(filename, converted_file)  # Create the zip file

        archive_path = converted_file[:-4]
        shutil.unpack_archive(converted_file, archive_path)  # Unzips to dir

        os.remove(converted_file)  # Get rid of the temp zip file

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
    soup = BeautifulSoup(workbook_xml_data, 'xml')

    # Find all <s:sheet> tags
    sheet_tags = soup.find_all('s:sheets')
    print(f'sheet tags: {sheet_tags}')

    # Extract name and sheetId values
    for sheet in sheet_tags:
        name = sheet.get('name')
        sheet_id = sheet.get('sheetId')
        print(f"Sheet Name: {name}, Sheet ID: {sheet_id}")


if __name__ == '__main__':
    file_to_read = r'test_data\test_sheet_01.xlsx'
    temp_archive_path = unzip_xlsx_file(file_to_read)
    get_sheets_info(temp_archive_path)
