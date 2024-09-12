from datetime import datetime, timedelta
from fractions import Fraction
import re


# Function systematically strips custom number formats down into their component parts so that they can be interpreted by python
def convert_custom_formats():  #(value, numFmt):
    value = '6516516516351.1'
    numFmt = '[Blue]_G£#,##0.00_);[Red]h:mm:ss AM/PM;dd/mm/yy;"sales "@" foo foo "'
    type = ''

    # Number formatting guidance from https://support.microsoft.com/en-gb/office/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5
    
    # ---STEP 1---
    # A number format can have up to four sections of code, separated by semicolons.
    # These code sections define the format for positive numbers, negative numbers, zero values, and text, in that order:
    # <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>
    # Detect semicolons and split:
    if ';' in str(numFmt):
        semicolon_split_numFmt = str(numFmt).split(';')
    else:
        semicolon_split_numFmt = []
        semicolon_split_numFmt.append(str(numFmt))

    # ---STEP 2---
    # You do not have to include all code sections in your custom number format.
    # If you specify only one code section, it is used for all numbers.
    # If you specify only two code sections for your custom number format, the first 
    # section is used for positive numbers and zeros, and the second section is used for negative numbers.
    if len(semicolon_split_numFmt) < 1:
        print(f'Invalid number format: {numFmt}')
        return value
    else:
        len(semicolon_split_numFmt)
        numFmt_dict = {'posNumFmt': {'numFmt_str': semicolon_split_numFmt[0]}}
        if len(semicolon_split_numFmt) >= 2:
            numFmt_dict['negNumFmt'] = {'numFmt_str': semicolon_split_numFmt[1]}
        if len(semicolon_split_numFmt) >= 3:
            numFmt_dict['zeroNumFmt'] = {'numFmt_str': semicolon_split_numFmt[2]}
        if len(semicolon_split_numFmt) == 4:
            numFmt_dict['textNumFmt'] = {'numFmt_str': semicolon_split_numFmt[3]}
        if len(semicolon_split_numFmt) > 4:
            print(f'Invalid number format: {numFmt}')

    for fmt in numFmt_dict:
        # print(numFmt_dict[fmt]['numFmt_str'])
        # TODO - check that numFmt_dict[fmt]['numFmt_str'] exists & isn't blank
        # TODO - there's a localisation tag in square brackets - filter for that before colours

        # ---STEP 3---
        # Check colours and list in dict, and strip from numFmt_str. N.B. Excel puts colours in square brackets
        # Should only be a single colour per numFmt
        colour = re.search(r'\[(.*?)\]', numFmt_dict[fmt]['numFmt_str'])
        if colour:
            numFmt_dict[fmt]['colour'] = colour.group(1)
            numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(f'[{colour.group(1)}]', '')

        # ---STEP 4---
        # Check padding/spacing and list in dict, and strip from numFmt_str.
        # N.B. Excel uses width of char after underscore to pad
        # Can potentially be multiple paddings in a single numFmt
        len_padded_numFmt_str = len(numFmt_dict[fmt]['numFmt_str'])
        paddings = re.finditer(r'_(.)', numFmt_dict[fmt]['numFmt_str'])
        if paddings:
            numFmt_dict[fmt]['paddings'] = []
            for padding in paddings:
                numFmt_dict[fmt]['paddings'].append([padding.group(1), padding.start(), len_padded_numFmt_str])  # Also returns position of padding in numFmt_str
        # Now strip padding from numFmt_str
        if len(numFmt_dict[fmt]['paddings']) > 0:
            for padding in numFmt_dict[fmt]['paddings']:
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(f'_{padding[0]}', '')
        else:
            del numFmt_dict[fmt]['paddings']  # Remove if empty

        # ---STEP 5---
        # Check for text strings and list in dict, and strip from numFmt_str. Similar to Step 4
        # N.B. Excel uses double quotes for text strings in numFmts
        # Can potentially be multiple text strings in a single numFmt
        len_text_strings_numFmt_str = len(numFmt_dict[fmt]['numFmt_str'])
        text_strings = re.finditer(r'"(.*?)"', numFmt_dict[fmt]['numFmt_str'])
        if text_strings:
            numFmt_dict[fmt]['text_strings'] = []
            for text_string in text_strings:
                numFmt_dict[fmt]['text_strings'].append([text_string.group(1), text_string.start(), len_text_strings_numFmt_str])
        # Now strip text_string from numFmt_str
        if len(numFmt_dict[fmt]['text_strings']) > 0:
            for text_string in numFmt_dict[fmt]['text_strings']:
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(f'"{text_string[0]}"', '')
        else:
            del numFmt_dict[fmt]['text_strings']  # Remove if empty

        # ---STEP 6---
        # Check for time formats, figure out their length, and infer whether mm is months or minutes
        if re.search(r'[dmyhs]', numFmt_dict[fmt]['numFmt_str']):
            print(f'Date format: {numFmt_dict[fmt]['numFmt_str']}')
            # First detect & strip AM/PM label
            if 'AM/PM' in numFmt_dict[fmt]['numFmt_str']:
                numFmt_dict[fmt]['AM_PM_label'] = True
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(' AM/PM', '')
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace('AM/PM', '')
            # Next, check formats for days, years, hours and seconds
            numFmt_dict[fmt]['d_count'] = numFmt_dict[fmt]['numFmt_str'].count('d')
            numFmt_dict[fmt]['y_count'] = numFmt_dict[fmt]['numFmt_str'].count('y')
            numFmt_dict[fmt]['h_count'] = numFmt_dict[fmt]['numFmt_str'].count('h')
            numFmt_dict[fmt]['s_count'] = numFmt_dict[fmt]['numFmt_str'].count('s')
            # TODO - figure out how to tell if mm is months or minutes
            # Remove empty counters
            time_counter_names = ['d_count', 'y_count', 'h_count', 's_count']
            for time_counter in time_counter_names:
                if numFmt_dict[fmt][time_counter] == 0:
                    del numFmt_dict[fmt][time_counter]
        
        # ---STEP 7---
        # numFmt strings can be wrapped in parentheses - common in accounting to mark negative numbers for visibility
        # Might as well strip those here
        parentheses = re.search(r'\((.*?)\)', numFmt_dict[fmt]['numFmt_str'])
        if parentheses:
            numFmt_dict[fmt]['parentheses'] = True
            numFmt_dict[fmt]['numFmt_str'] = parentheses.group(1)
        
        # ---STEP 8---
        # Detect and strip currency symbols.
        # Assume there's only going to be one per numFmt as it's invalid otherwise
        currency_pattern = r'[\$\€\£\¥\₹\₽\₩\₪]'  # Add additional currency symbols here if needed
        currencies = re.search(currency_pattern, numFmt_dict[fmt]['numFmt_str'])
        if currencies:
            numFmt_dict[fmt]['currency'] = currencies.group()
            numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(currencies.group(), '')
        
        # ---STEP 9---
        # Detect and strip negative symbols.
        # Assume there's only going to be one per numFmt, & in position [0], as it's invalid otherwise
        if len(numFmt_dict[fmt]['numFmt_str']) > 0:  # numFmt can have zero length if there's no entry
            if numFmt_dict[fmt]['numFmt_str'][0] == '-':
                numFmt_dict[fmt]['negative'] = True
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'][1:]
        
        # ---STEP 10---
        # Detect decimal point if present, and split numFmt, to give number of decimal places
        numFmt_decimal_split = numFmt_dict[fmt]['numFmt_str'].split('.')
        if len(numFmt_decimal_split) < 1:
            print(f'Invalid number format: {numFmt}')
        elif len(numFmt_decimal_split) == 1:
            pass
        elif len(numFmt_decimal_split) == 2:
            numFmt_dict[fmt]['numFmt_str'] = numFmt_decimal_split[0]
            numFmt_dict[fmt]['decimal_places'] = len((numFmt_decimal_split[1]))
        else:  # Valid numbers shouldn't have more than one decimal place
            print(f'Invalid number format: {numFmt}')
        
        # ---STEP 11---
        # Detect and strip leading hashes indicating thousands separators or similar
        if len(numFmt_dict[fmt]['numFmt_str']) > 0:  # numFmt can have zero length if there's no entry
            if numFmt_dict[fmt]['numFmt_str'][0] == '#':
                # Regular expression to find hash pattern
                numFmt_dict[fmt]['hashes'] = re.findall(r'\D+', numFmt_dict[fmt]['numFmt_str'])[0]  # Find sequences of non-digits
                numFmt_dict[fmt]['numFmt_str'] = numFmt_dict[fmt]['numFmt_str'].replace(numFmt_dict[fmt]['hashes'], '')


    for number_format in numFmt_dict:
        print(f'{number_format}: {numFmt_dict[number_format]}')
    

    # # Check if int or float
    # try:
    #     int_value = int(value)
    #     type = 'int'
    # except ValueError:
    #     # If it raises a ValueError, it's not an int
    #     pass
    
    # # Try to convert to float
    # try:
    #     float_value = float(value)
    #     type = 'float'
    # except ValueError:
    #     # If it raises a ValueError, it's not a float
    #     pass

    # if '.' in str(numFmt):
    #     decimal_split_numFmt = str(numFmt).split('.')  # TODO - this is only going to work for formats with exactly one '.'
    #     print(f'decimal, {type}, {len(decimal_split_numFmt[1])}d.p.')
    # else:
    #     print(f'int? {type}')


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
        return date_value.strftime('%d-%b')  # Format as "DD-mmm"

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

    # formatted_number = apply_excel_numFmtId(test_value, num_format_reference)
    
    # print(f'{formatted_number}')
    convert_custom_formats()

    print('   --- END ---')
