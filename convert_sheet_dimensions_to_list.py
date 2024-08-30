
def dimensions_to_list(dims):
    start_ref = dims.split(':')[0]
    end_ref = dims.split(':')[1]

    print(start_ref, end_ref)
    column_letter = 'AA63'[0]
    row_number = int('AA63'[1:])
    column_index = ord(column_letter.upper()) - ord('A')
    print(column_letter, row_number, column_index)



if __name__ == '__main__':
    print('   --- START ---')

    dimensions = 'A1:B31'
    dim_list = dimensions_to_list(dimensions)
    print(dim_list)

    print('   --- END ---')