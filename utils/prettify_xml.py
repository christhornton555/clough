# Utility to prettify XML files to make them easier to read

import xml.etree.ElementTree as ET
import xml.dom.minidom
from io import StringIO

def pretty_print_xml(input_file, output_file):
    # Parse the XML file
    tree = ET.parse(input_file)
    root = tree.getroot()
    
    # Convert the ElementTree to a string
    xml_str = ET.tostring(root)

    # Use StringIO to create an in-memory file
    in_mem_file = StringIO()
    in_mem_file.write('<?xml version="1.0" encoding="UTF-8"?>\n')  # XML declaration
    in_mem_file.write(xml_str.decode('utf-8'))  # Write the actual XML content
    in_mem_file.seek(0)  # Reset the cursor position

    # Parse the in-memory XML string using minidom
    dom = xml.dom.minidom.parseString(in_mem_file.read())
    pretty_xml_as_string = dom.toprettyxml()
    
    # Write the pretty XML to the output file
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(pretty_xml_as_string)

if __name__ == '__main__':
    print('   --- START ---')

    # Just put the relative filepath here
    input_file = r'test_data/test_sheet_01 - Copy/xl/styles.xml'
    output_file = input_file[:-4] + '_PRETTIFIED.xml'
    pretty_print_xml(input_file, output_file)
    
    print('   --- END ---')
