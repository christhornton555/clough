'''
This script takes a number format reference and finds and returns the number format it relates to
'''

def make_num_formats_dict():
    standard_num_formats_file_path = 'assets/numFmtIds.csv'
    standard_num_formats_list = [] 
    with open(standard_num_formats_file_path, "r") as file:
        for line in file:
            # Remove any leading/trailing whitespace and append to the list
            standard_num_formats_list.append(line.strip().split('|'))

    standard_num_formats_dict = {}
    for item in standard_num_formats_list:
        key, value = int(item[0]), item[1]
        standard_num_formats_dict[key] = value

    return standard_num_formats_dict

if __name__ == '__main__':
    print('   --- START ---')

    num_format_reference = 15
    temp_archive_path = 'test_data/output/20240829_000803_output'

    standard_num_formats_dict = make_num_formats_dict()
    
    print(standard_num_formats_dict[num_format_reference])

    print('   --- END ---')