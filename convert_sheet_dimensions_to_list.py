
def split_cell_reference_into_rows_and_columns(cell_reference):
    for char in range(len(cell_reference)):
        if not cell_reference[char].isalpha():
            split_start_ref_at = char
            break
    return [cell_reference[:split_start_ref_at], cell_reference[split_start_ref_at:]]


def dimensions_to_list(dims):
    start_ref = dims.split(':')[0]
    end_ref = dims.split(':')[1]

    return[split_cell_reference_into_rows_and_columns(start_ref), split_cell_reference_into_rows_and_columns(end_ref)]


if __name__ == '__main__':
    print('   --- START ---')

    dimensions = 'ABC123:BDGJ31457'
    dim_list = dimensions_to_list(dimensions)
    print(dim_list)

    print('   --- END ---')