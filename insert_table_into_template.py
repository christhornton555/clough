# from datetime import datetime
from get_raw_data_from_xlsx_sheets import get_current_time_string  # I should probably put this in its own module

TABLE_INSERTION_POINT = '<!-- #CLOUGH -->'  # The key string which will be replaced with the table data from Excel


def split_template(template_file_path):
    with open(template_file_path, 'r') as template_file:
        template_html = template_file.read()

    # Split the template_html into two parts using the split method
    template_top, template_tail = template_html.split(TABLE_INSERTION_POINT, 1)
    # The `1` in the split method ensures that the string is split at the first occurrence of the marker

    return template_top, template_tail


def insert_table_into_template(template_file, table_strings):
    current_time = get_current_time_string()
    output_file_path = f'test_data/output/{current_time}_cloughtest_output.html'
    for sheet in table_strings:  # TODO - make this handle multisheets
        table_to_insert = table_strings[sheet]
        html_template_top, html_template_tail = split_template(template_file)
        full_html = html_template_top + table_to_insert + html_template_tail

    with open(output_file_path, 'w') as output_file:
        output_file.write(full_html)

        output_file.close()



if __name__ == '__main__':
    print('   --- START ---')
    
    html_template_file = r'test_data/cloughtest_template.html'
    all_table_strings = {'Sheet1': '<table class="clough" id="Sheet1" style="width:100%">\n\t<caption>Converted from Excel by Clough</caption>\n\t<thead>\n\t\t<tr>\n\t\t\t<th></th>\n\t\t\t<th>A</th>\n\t\t\t<th>B</th>\n\t\t</tr>\n\t</thead>\n\n\t<tbody>\n\t\t<tr>\n\t\t\t<td>1</td>\n\t\t\t<td>Date</td>\n\t\t\t<td>Rainfall (mm)</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>2</td>\n\t\t\t<td>01-Aug-24</td>\n\t\t\t<td>5.2</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>3</td>\n\t\t\t<td>02-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>4</td>\n\t\t\t<td>03-Aug-24</td>\n\t\t\t<td>12.4</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>5</td>\n\t\t\t<td>04-Aug-24</td>\n\t\t\t<td>3.8</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>6</td>\n\t\t\t<td>05-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>7</td>\n\t\t\t<td>06-Aug-24</td>\n\t\t\t<td>18.7</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>8</td>\n\t\t\t<td>07-Aug-24</td>\n\t\t\t<td>6.3</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>9</td>\n\t\t\t<td>08-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>10</td>\n\t\t\t<td>09-Aug-24</td>\n\t\t\t<td>2.9</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>11</td>\n\t\t\t<td>10-Aug-24</td>\n\t\t\t<td>14.2</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>12</td>\n\t\t\t<td>11-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>13</td>\n\t\t\t<td>12-Aug-24</td>\n\t\t\t<td>4.5999999999999996</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>14</td>\n\t\t\t<td>13-Aug-24</td>\n\t\t\t<td>9.1</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>15</td>\n\t\t\t<td>14-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>16</td>\n\t\t\t<td>15-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>17</td>\n\t\t\t<td>16-Aug-24</td>\n\t\t\t<td>23.3</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>18</td>\n\t\t\t<td>17-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>19</td>\n\t\t\t<td>18-Aug-24</td>\n\t\t\t<td>1.5</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>20</td>\n\t\t\t<td>19-Aug-24</td>\n\t\t\t<td>7.8</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>21</td>\n\t\t\t<td>20-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>22</td>\n\t\t\t<td>21-Aug-24</td>\n\t\t\t<td>15.6</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>23</td>\n\t\t\t<td>22-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>24</td>\n\t\t\t<td>23-Aug-24</td>\n\t\t\t<td>8.1999999999999993</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>25</td>\n\t\t\t<td>24-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>26</td>\n\t\t\t<td>25-Aug-24</td>\n\t\t\t<td>13.7</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>27</td>\n\t\t\t<td>26-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>28</td>\n\t\t\t<td>27-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>29</td>\n\t\t\t<td>28-Aug-24</td>\n\t\t\t<td>19.399999999999999</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>30</td>\n\t\t\t<td>29-Aug-24</td>\n\t\t\t<td>0</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td>31</td>\n\t\t\t<td>30-Aug-24</td>\n\t\t\t<td>11.3</td>\n\t\t</tr>\n\t</tbody>\n\n\t<tfoot>\n\t</tfoot>\n\n</table>\n'}

    insert_table_into_template(html_template_file, all_table_strings)

    print('   --- END ---')