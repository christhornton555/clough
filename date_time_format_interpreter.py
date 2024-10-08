# Test file to figure out the flow control for interpreting custom datetime formats
# The script apply_number_formatting.py is getting messy, so let's figure it out here

import re

dt_formats = ['myhm','yyymms mmm','ddmmyy hhmmmss','yyymm', 'yyyymm - s mmm', 'yyhhmm - s mmm ddmm', 'ddhhyyy', 'mmmmdddd']

for i in range(len(dt_formats)):
    if dt_formats[i].count('m') > 0:  # Are there any 'm's present? TODO - what happens with upper case m's?
        m_sequences = re.findall(r'm+', dt_formats[i])
        print(f'{i}. Sequences found: {m_sequences}')
    else:
        print(f'{i}. No sequences found')
