# Utility to prettify XML files to make them easier to read

import xml.etree.ElementTree as ET
import xml.dom.minidom

def pretty_print_xml(input_file, output_file):
    # Parse the XML file
    tree = ET.parse(input_file)
    root = tree.getroot()
    
    # Convert the ElementTree to a string
    xml_str = ET.tostring(root, encoding='unicode')

    # Use minidom to pretty print the XML string
    dom = xml.dom.minidom.parseString(xml_str)
    pretty_xml_as_string = dom.toprettyxml()
    
    # Write the pretty XML to the output file
    with open(output_file, 'w') as file:
        file.write(pretty_xml_as_string)

# Just put the relative filepath here
input_file = r'test_data\test_sheet_01 - Copy\xl\theme\theme1.xml'
output_file = input_file[:-4] + '_PRETTIFIED.xml'
pretty_print_xml(input_file, output_file)
