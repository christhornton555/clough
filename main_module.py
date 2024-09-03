'''
This script calls all the other modules
'''

from make_workbook_dict import make_workbook_dict
from convert_data_to_html import convert_data_to_html
from insert_table_into_template import insert_table_into_template

if __name__ == '__main__':
    print('   --- START ---')

    file_to_read = r'test_data/test_sheet_01.xlsx'
    html_template_file = r'test_data/cloughtest_template.html'

    workbook_to_output = make_workbook_dict(file_to_read)

    all_table_strings = convert_data_to_html(workbook_to_output)

    insert_table_into_template(html_template_file, all_table_strings)

    print('   --- END ---')