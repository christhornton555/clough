'''
This script takes a style reference and find and returns the style it relates to
'''
from bs4 import BeautifulSoup  # For picking data out of an XML file 

def get_style(ref, archive_path):
    styles_file_path = f'{archive_path}/xl/styles.xml'
    with open(styles_file_path, 'r') as styles_file:
        styles_xml_data = styles_file.read()

    # Create a BeautifulSoup object
    styles_soup = BeautifulSoup(styles_xml_data, features='xml')

    cell_xfs_tags = styles_soup.find('cellXfs')  # Reads all defined cell styles into list
    style_tags = cell_xfs_tags.find_all('xf')
    alignment_tags = cell_xfs_tags.find_all('alignment')

    # for tag in alignment_tags:
    #     print(tag)

    return style_tags[ref].get('numFmtId')


if __name__ == '__main__':
    print('   --- START ---')

    style_reference = 2
    temp_archive_path = 'test_data/output/20240829_000803_output'
    print(get_style(style_reference, temp_archive_path))

    print('   --- END ---')