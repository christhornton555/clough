'''
This script takes a style reference and find and returns the style it relates to
'''
from bs4 import BeautifulSoup  # For picking data out of an XML file 

def get_font(font_ref, styles):
    font_tags = styles.find_all('font')
    return font_tags[int(font_ref)]

def get_style(style_ref, archive_path):
    styles_file_path = f'{archive_path}/xl/styles.xml'
    with open(styles_file_path, 'r') as styles_file:
        styles_xml_data = styles_file.read()

    # Create a BeautifulSoup object
    styles_soup = BeautifulSoup(styles_xml_data, features='xml')

    cell_xfs_tags = styles_soup.find('cellXfs')  # Reads all defined cell styles into list
    style_tags = cell_xfs_tags.find_all('xf')
    alignment_tags = cell_xfs_tags.find_all('alignment')

    # Now convert the fontId value in this style into actual details about styling
    font_id = style_tags[style_ref].get('fontId')
    font = get_font(font_id, styles_soup)

    output_dict = {
        'numFmtId': style_tags[style_ref].get('numFmtId')
    }
    if int(output_dict['numFmtId']) > 49:
        print(f'Full style: {style_tags[style_ref]}')

    
    try:
        if alignment_tags[style_ref-1].get('horizontal') != None:
            output_dict['horizontal_alignment'] = alignment_tags[style_ref-1].get('horizontal')  # 1st (default) <xf> tag has no <alignment> sub-tag, hence ref-1
    except IndexError:
        print(f'Index error - {alignment_tags}, {len(alignment_tags)}, {style_ref-1}')

    
    # Add font styles (bold, italic, underline) if present
    output_dict['font_style'] = ''
    if font.find('b') != None:
        output_dict['font_style'] += 'bold '
    
    if font.find('i') != None:
        output_dict['font_style'] += 'italic '
    
    if font.find('u') != None:
        output_dict['font_style'] += 'underline '

    # Remove the 'font_style' key if it's empty
    if not output_dict['font_style']:
        output_dict.pop('font_style', None)  # Remove key if empty

    return output_dict


if __name__ == '__main__':
    print('   --- START ---')

    style_reference = 3
    temp_archive_path = 'test_data/output/20240829_000803_output'
    print(get_style(style_reference, temp_archive_path))

    print('   --- END ---')