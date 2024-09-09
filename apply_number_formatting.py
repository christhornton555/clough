from datetime import datetime, timedelta
from fractions import Fraction

def apply_excel_numFmtId(value, numFmtId):
    # TODO - add a type-check to convert value to int/floats if necessary
    # TODO - gonna need some kind of localisation check, expecially for dates & currency, as Excel evidently handles that automatically

    numFmtId = int(numFmtId)

    # Several formats use dates, so do the core conversion here:
    base_date = datetime(1899, 12, 30)  # Excel uses 1900-based system (-1, cos they got their leap years wrong)
    date_value = base_date + timedelta(days=int(value))

    if numFmtId == 0:  # General number format (just return as is)
        if isinstance(value, float):
            # For float values, return as is but remove trailing decimal places if not needed
            return str(value).rstrip('0').rstrip('.') if '.' in str(value) else str(value)
        else:
            # For other types (like int), just return the value as a string
            return str(value)

    elif numFmtId == 1:  # Format = 0
        return f"{int(value)}"

    elif numFmtId == 2:  # Format = 0.00
        return f"{value:.2f}"

    elif numFmtId == 3:  # Format = #,##0
        return f"{value:,.0f}"

    elif numFmtId == 4:  # Format = #,##0.00
        return f"{value:,.2f}"

    elif numFmtId == 9:  # Format = 0%
        return f"{value * 100:.0f}%"

    elif numFmtId == 10:  # Format = 0.00%
        return f"{value * 100:.2f}%"

    elif numFmtId == 11:  # Format = 0.00E+00
        return "{:.2E}".format(value)  # 2 decimal places in scientific notation

    elif numFmtId == 12:  # Format = # ?/?
        # Convert the float value to a fraction with limit_denominator for precision
        fraction_value = Fraction(value).limit_denominator(100)  # Limit to denominators of 100 or less
        return f"{fraction_value}"

    elif numFmtId == 13:  # Format = # ??/??
        # Convert the value to a fraction with a denominator up to 99 (to match ??/?? format)
        fraction_value = Fraction(value).limit_denominator(99)
        # Format the fraction as "numerator/denominator"
        return f"{fraction_value.numerator} {fraction_value.denominator}"

    elif numFmtId == 14:  # Date format (typically short date in Excel)
        # Assuming `value` is a float representing Excel date (days since 1900)
        # TODO - add in a way to handle dates before 02-01-1900
        return date_value.strftime('%m-%d-%y')  # Format to MM-DD-YY
    
    elif numFmtId == 15:  # Date format (typically short date in Excel)
        # Assuming `value` is a float representing Excel date (days since 1900)
        # TODO - add in a way to handle dates before 02-01-1900
        return date_value.strftime('%d-%b-%y')  # Format as "DD-mmm-YYY"
    
    elif numFmtId == 16:  # Date format (typically short date in Excel)
        # Assuming `value` is a float representing Excel date (days since 1900)
        # TODO - add in a way to handle dates before 02-01-1900
        return date_value.strftime('%d-%b')  # Format as "DD-mmm-YYY"

    elif numFmtId == 21:  # Time format
        hours = int(value * 24)
        minutes = int((value * 24 * 60) % 60)
        seconds = int((value * 24 * 3600) % 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    elif numFmtId == 165:  # Format = DD/MM/YY
        return date_value.strftime('%d/%m/%y')  # Format to DD/MM/YY

    else:
        print(f"Unsupported numFmtId: {numFmtId}")
        return f"{value}"



if __name__ == '__main__':
    print('   --- START ---')

    num_format_reference = 11
    test_value = 2

    formatted_number = apply_excel_numFmtId(test_value, num_format_reference)
    
    print(f'{formatted_number}')

    print('   --- END ---')
