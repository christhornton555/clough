'''
This script calls all the other modules
'''

from make_workbook_dict import make_workbook_dict
from convert_data_to_html import convert_data_to_html

if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'
    workbook_to_output = make_workbook_dict(file_to_read)
    convert_data_to_html(workbook_to_output)

    print('   --- END ---')