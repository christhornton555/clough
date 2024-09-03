'''
This script takes a dictionary containing a worksheet's data, and produces an HTML table of it
'''

def convert_data_to_html(table_data):
    # print(data)
    table_style = {
        'width': '100%',
        'caption': 'Converted from Excel by Clough'
    }

    table_headers = table_data['Sheet1'][0]  # TODO - Set 'Sheet1' programmatically
    table_headers_string = ''
    for table_header in table_headers:
        table_headers_string += f'\t\t<th>{table_header}</th>\n'

    full_table_string = (
        f'<table style="width:{table_style['width']}">\n'
        f'\t<caption>{table_style['caption']}</caption>\n'
        f'\t<tr>\n'
        f'{table_headers_string}'
        f'\t</tr>\n'
        )

    print(full_table_string)


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
    
    convert_data_to_html(test_data)

    print('   --- END ---')