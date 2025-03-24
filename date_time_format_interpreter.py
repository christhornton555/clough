import re
import datetime



def split_excel_format(format_string):
    # Split the format string into tokens
    tokens = []
    i = 0
    length = len(format_string)

    while i < length:
        if format_string[i] == '\\' and i + 1 < length:
            # Escape sequence: \x → literal x
            tokens.append(format_string[i+1])
            i += 2
        elif format_string[i].isalpha():
            # Start of a format token
            j = i + 1
            while j < length and format_string[j] == format_string[i]:
                j += 1
            token = format_string[i:j]
            tokens.append(token)
            i = j
        else:
            # Start of punctuation or space
            j = i + 1
            while j < length and not format_string[j].isalpha() and format_string[j] != '\\':
                j += 1
            tokens.append(format_string[i:j])
            i = j

    return tokens

def interpret_mins_or_month(full_datetime_format):
    print('"m" found')

def format_datetime(raw_datetime_value, excel_datetime_format):
    # Excel considers 1900-01-01 as day 1, but Python's datetime starts at 1900-01-01 as day 0
    # Excel incorrectly treats 1900 as a leap year, so we subtract 2 days to align
    base_date = datetime.datetime(1899, 12, 30)
    days = int(raw_datetime_value)
    fractional_day = raw_datetime_value - days
    seconds = round(fractional_day * 86400)  # Round like this to avoid discrepancies with rounding errors
    delta = datetime.timedelta(days=days, seconds=seconds)

    converted_datetime = base_date + delta  # Excel datetime float is now a Python datetime object. Now we need to format it


    python_datetime_format = ''  # We'll concatenate Python datetime formatting onto this string
    excel_datetime_format = excel_datetime_format.lower()  # Excel datetime formats are case insensitive

    tokenised_excel_datetime_format = split_excel_format(excel_datetime_format)

    print(tokenised_excel_datetime_format)
    

    if excel_datetime_format.count('m') > 0:
        interpret_mins_or_month(excel_datetime_format)

    else:
        print('"m" not found')

    python_datetime_format += '%d %B %Y %H:%M:%S'
    
    return converted_datetime.strftime(python_datetime_format)



sample_time = 28714.5068981481  # 12/08/1978 12:09:56

dt_formats = ['yyyymmdd hh:mm:ss']

for i in range(len(dt_formats)):
    formatted_datetime = format_datetime(sample_time, dt_formats[i])

    print(formatted_datetime)
