'''
This script takes a string reference and find and returns the string it relates to
'''
from bs4 import BeautifulSoup  # For picking data out of an XML file 

def get_string(ref, archive_path):
    strings_file_path = f'{archive_path}/xl/sharedStrings.xml'
    with open(strings_file_path, 'r') as strings_file:
        strings_xml_data = strings_file.read()

    # Create a BeautifulSoup object
    strings_soup = BeautifulSoup(strings_xml_data, features='xml')

    string_tags = strings_soup.find_all('si')

    print(string_tags[ref])

    # print(string_tags[ref].find('t').text)

    return string_tags[ref].find('t').text


if __name__ == '__main__':
    print('   --- START ---')

    string_reference = 0
    temp_archive_path = 'test_data/output/20240829_000803_output'
    get_string(string_reference, temp_archive_path)

    print('   --- END ---')