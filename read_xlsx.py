'''
This script opens a basic .xlsx file, reads its contents, and displays them to screen
'''
import shutil  # Various utilities including for copying files
import os  # Handles things like deleting files
from datetime import datetime  # For timestamping files


# Returns the current time as a string for timestamping filenames
# TODO - I might move this to its own module - I anticipate using it a lot
def get_current_time_string():
    now = datetime.now()
    formatted_time = now.strftime('%Y%m%d_%H%M%S')
    return formatted_time


# Copies any xlsx file to a zip, which it then extracts
def unzip_xlsx_file(filename):
    if filename[-5:] == '.xlsx':
        current_time = get_current_time_string()
        converted_file = f'test_data/{current_time}_output.zip'
        shutil.copy(filename, converted_file)  # Create the zip file

        archive_name = converted_file[:-4]
        shutil.unpack_archive(converted_file, archive_name)  # Unzips to dir

        os.remove(converted_file)  # Get rid of the temp zip file

        print(f'File unzipped to path {archive_name}')

    else:
        print(f'Specified file, {filename}, was not a .xsls file.')



if __name__ == '__main__':
    file_to_read = r'test_data\test_sheet_01.xlsx'
    unzip_xlsx_file(file_to_read)
