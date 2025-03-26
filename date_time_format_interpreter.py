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
            tokens.append('\\' + format_string[i+1])
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


def find_positions_of_m_tokens(tokenised_excel_datetime_format):
    valid_m_positions = []

    for i in range(len(tokenised_excel_datetime_format)):
        # "m" is in the first token
        if tokenised_excel_datetime_format[i].count('m') > 0 and i == 0:
            # print(f'i == {i}, m_token = {tokenised_excel_datetime_format[i]}, next token = {tokenised_excel_datetime_format[i+1]}')
            valid_m_positions.append(i)

        # 'm' is in a token somewhere in the middle
        elif tokenised_excel_datetime_format[i].count('m') > 0 and i > 0 and i < len(tokenised_excel_datetime_format) - 1:
            # print(f'i == {i} (>0), m_token = {tokenised_excel_datetime_format[i]}, previous token = {tokenised_excel_datetime_format[i-1]}, next token = {tokenised_excel_datetime_format[i+1]}')
            valid_m_positions.append(i)

        # 'm' is in the last token
        elif tokenised_excel_datetime_format[i].count('m') > 0 and i == len(tokenised_excel_datetime_format) - 1:
            # print(f'i == {i} (len), m_token = {tokenised_excel_datetime_format[i]}, previous token = {tokenised_excel_datetime_format[i-1]}')
            valid_m_positions.append(i)

        # 'm' is not in this token
        elif tokenised_excel_datetime_format[i].count('m') == 0:
            # print(f'i == {i}, no "m" in this token')
            pass

        else:  # Should be unreachable
            print(f'Error. i == {i}, m_token = {tokenised_excel_datetime_format[i]}')

    return valid_m_positions


def interpret_mins_or_month(tokenised_excel_datetime_format):
    positions = find_positions_of_m_tokens(tokenised_excel_datetime_format)
    print(f'"m" found at {positions}')
    minute_token_positions = set()

    for pos in positions:
        current_token = tokenised_excel_datetime_format[pos]
        is_minute = False

        # Look backwards
        i = pos - 1
        while i >= 0:
            token = tokenised_excel_datetime_format[i]
            if token.isalpha():
                if 'h' in token:
                    is_minute = True
                    break
                elif 's' in token:
                    is_minute = True
                    break
                else:
                    break  # found an unrelated letter token
            i -= 1

        # Look forwards only if not already decided
        if not is_minute:
            i = pos + 1
            while i < len(tokenised_excel_datetime_format):
                token = tokenised_excel_datetime_format[i]
                if token.isalpha():
                    if 's' in token:
                        is_minute = True
                    break  # found a letter token, whether it's 's' or not
                i += 1

        if is_minute:
            minute_token_positions.add(pos)

        interpretation = 'minute' if is_minute else 'month'
        print(f'Token "{current_token}" at position {pos} is interpreted as: {interpretation}')

    return minute_token_positions
    

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

    format_conversion_lookup_dict = {
        # Year
        'yyyy': '%Y',  # Four-digit year
        'yyy': '%Y',   # Not standard in Excel, but map to four-digit year
        'yy': '%y',    # Two-digit year
        'y': '%y',     # Not standard, fallback to 2-digit year

        # N.B. Only defining the "m" values as months. There's separate logic to handle minutes
        # Month
        'mmmm': '%B',  # Full month name
        'mmm': '%b',   # Abbreviated month name
        'mm': '%m',    # Two-digit month number (01–12)
        'm': '%-m',    # One-digit month number (1–12)

        # Day
        'dddd': '%A',  # Full weekday name
        'ddd': '%a',   # Abbreviated weekday name
        'dd': '%d',    # Two-digit day of month (01–31)
        'd': '%-d',    # One-digit day of month (1–31)

        # Hour
        'hh': '%H',    # Two-digit hour (00–23)
        'h': '%-H',    # One-digit hour (0–23)

        # Second
        'ss': '%S',    # Two-digit seconds (00–59)
        's': '%-S',    # One-digit seconds (0–59)

        # AM/PM
        'am/pm': '%p',  # AM or PM
        'a/p': '%p'     # Excel shorthand (A/P) → AM/PM in Python
    }
    

    minute_positions = set()
    if excel_datetime_format.count('m') > 0:
        minute_positions = interpret_mins_or_month(tokenised_excel_datetime_format)

    else:
        print('"m" not found')
    
    for idx, token in enumerate(tokenised_excel_datetime_format):
        if token in ['m', 'mm'] and idx in minute_positions:
            # Disambiguated as minutes
            python_datetime_format += '%M' if token == 'mm' else '%-M'
        elif token in format_conversion_lookup_dict:
            python_datetime_format += format_conversion_lookup_dict[token]
        elif token.startswith('\\'):  # Handle escaped characters
            python_datetime_format += token[1]
        else:
            python_datetime_format += token  # Add punctuation, excaped characters etc, unchanged
    
    return converted_datetime.strftime(python_datetime_format)



sample_time = 28714.5068981481  # 12/08/1978 12:09:56

dt_formats = ['yyyy/mm/dd hh:mm:ss \\s']  # TODO - make sure escaped characters get double-escaped

for i in range(len(dt_formats)):
    formatted_datetime = format_datetime(sample_time, dt_formats[i])

    print(formatted_datetime)
