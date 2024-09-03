'''
This script calls all the other modules
'''

from datetime import datetime
from make_table import make_table

if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'
    workbook_to_output = make_table(file_to_read)
    print(workbook_to_output)

    print('   --- END ---')