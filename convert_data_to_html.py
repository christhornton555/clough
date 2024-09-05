'''
This script takes a dictionary containing a worksheet's data, and produces an HTML table of it
'''

def convert_data_to_html(table_data):
    table_style = {
        'width': '100%',
        'caption': 'Converted from Excel by Clough'
    }

    all_table_strings = {}
    css_class_definitions_to_add = []  # Classes will be added to this list as they're detected, then written into <head><style> later

    for sheet in table_data:
        sheet_columns = len(max(table_data[sheet], key=len))  # Find the longest row length, in case of rows of differing length

        # Define classes to later format as necessary
        table_classes_string = ''
        table_caption_classes_string = ''
        table_head_classes_string = ''
        table_body_classes_string = ''
        table_foot_classes_string = ''
        table_th_classes_string = ''
        table_tr_classes_string = ''
        table_td_classes_string = ''

        table_headers = table_data[sheet][0]
        table_headers_string = ''
        for table_header in table_headers:
            table_headers_string += f'\t\t\t<th class="{table_th_classes_string}">{table_header}</th>\n'

        table_body_string = ''
        for row in range(1, len(table_data[sheet])):  # Skip the first row which just has column labels A, B, C, etc
            table_body_string += f'\t\t<tr class="{table_tr_classes_string}">\n'
            for col in range(sheet_columns):
                cell_metadata = ''
                if col > 0:  # Ignore first column, which just has row numbers
                    cell_metadata = table_data[sheet][row][col][1]
                    if cell_metadata['horizontal_alignment'] != None:
                        table_td_classes_string = f'clough-align-{cell_metadata['horizontal_alignment']}'
                    else:
                        table_td_classes_string = ''

                    table_body_string += f'\t\t\t<td class="{table_td_classes_string}">{table_data[sheet][row][col][0]}</td>\n'
                else:
                    table_body_string += f'\t\t\t<td class="{table_td_classes_string}">{table_data[sheet][row][col]}</td>\n'
                table_td_classes_string = ''  # Clear last values
            table_body_string += f'\t\t</tr>\n'

        table_foot_string = ''  # Not currently implemented - included for completeness

        full_table_string = (
            f'<table class="clough {table_classes_string}" id="{sheet}" style="width:{table_style['width']}">\n'
            f'\t<caption class="{table_caption_classes_string}">{table_style['caption']}</caption>\n'
            f'\t<thead class="{table_head_classes_string}">\n'
            f'\t\t<tr class="{table_tr_classes_string}">\n'
            f'{table_headers_string}'
            f'\t\t</tr>\n'
            f'\t</thead>\n\n'

            f'\t<tbody class="{table_body_classes_string}">\n'
            f'{table_body_string}'
            f'\t</tbody>\n\n'

            f'\t<tfoot class="{table_foot_classes_string}">\n'
            f'{table_foot_string}'
            f'\t</tfoot>\n\n'
            f'</table>\n'
            )

        all_table_strings[sheet] = full_table_string
    return all_table_strings, css_class_definitions_to_add


if __name__ == '__main__':
    print('   --- START ---')

    test_data = {'Sheet1': [['', 'A', 'B'], 
                            [1, 'Date', 'Rainfall (mm)'], 
                            [2, '01-Aug-24', '5.2'], 
                            [3, '02-Aug-24', '0'], 
                            [4, '03-Aug-24', '12.4'], 
                            [5, '04-Aug-24', '3.8'], 
                            [6, '05-Aug-24', '0'], 
                            [7, '06-Aug-24', '18.7'], 
                            [8, '07-Aug-24', '6.3'], 
                            [9, '08-Aug-24', '0'], 
                            [10, '09-Aug-24', '2.9'], 
                            [11, '10-Aug-24', '14.2'], 
                            [12, '11-Aug-24', '0'], 
                            [13, '12-Aug-24', '4.5999999999999996'], 
                            [14, '13-Aug-24', '9.1'], 
                            [15, '14-Aug-24', '0'], 
                            [16, '15-Aug-24', '0'], 
                            [17, '16-Aug-24', '23.3'], 
                            [18, '17-Aug-24', '0'], 
                            [19, '18-Aug-24', '1.5'], 
                            [20, '19-Aug-24', '7.8'], 
                            [21, '20-Aug-24', '0'], 
                            [22, '21-Aug-24', '15.6'], 
                            [23, '22-Aug-24', '0'], 
                            [24, '23-Aug-24', '8.1999999999999993'], 
                            [25, '24-Aug-24', '0'], 
                            [26, '25-Aug-24', '13.7'], 
                            [27, '26-Aug-24', '0'], 
                            [28, '27-Aug-24', '0'], 
                            [29, '28-Aug-24', '19.399999999999999'], 
                            [30, '29-Aug-24', '0'], 
                            [31, '30-Aug-24', '11.3']]}
    
    table_strings, css = convert_data_to_html(test_data)
    # print(css)

    print('   --- END ---')