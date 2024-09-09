from datetime import datetime, timedelta

def apply_excel_numFmtId(value, numFmtId):
    # TODO - add a type-check to convert inputs to ints/floats if necessary

    if numFmtId == 1:  # General number format (just return as is)
        return f"{value}"

    elif numFmtId == 10:  # Percentage format
        return f"{value * 100:.2f}%"

    elif numFmtId == 14:  # Date format (typically short date in Excel)
        # Assuming `value` is a float representing Excel date (days since 1900)
        base_date = datetime(1899, 12, 30)  # Excel uses 1900-based system
        date_value = base_date + timedelta(days=value)
        return date_value.strftime('%m/%d/%Y')  # Format to MM/DD/YYYY

    elif numFmtId == 21:  # Time format
        hours = int(value * 24)
        minutes = int((value * 24 * 60) % 60)
        seconds = int((value * 24 * 3600) % 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    else:
        print(f"Unsupported numFmtId: {numFmtId}")
        return f"{value}"




if __name__ == '__main__':
    print('   --- START ---')

    num_format_reference = 14
    test_value = 0.8

    formatted_number = apply_excel_numFmtId(test_value, num_format_reference)
    
    print(f'{formatted_number}')

    print('   --- END ---')
