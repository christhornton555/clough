# Test file to figure out the flow control for interpreting custom datetime formats
# The script apply_number_formatting.py is getting messy, so let's figure it out here

import re

dt_formats = ['myhm','yyymms mmm','ddmmyy hhmmmss','yyymm', 'yyyymm - s mmm', 'yyhhmm - s mmm ddmm', 'ddhhyyy', 'mmmmdddd']

output_list = []

for i in range(len(dt_formats)):
    dt_format = dt_formats[i].lower()  # Excel is case insensitive, so just convert all to lower
    if dt_format.count('m') > 0:  # Are there any 'm's present? TODO - what happens with upper case m's?
        m_sequences = re.findall(r'm+', dt_format)

        seq_analysis_list = []
        for seq in m_sequences:  # Sort out all the different sequence lengths

            # Could be mins or months
            if len(seq) == 1:
                seq_analysis_list.append(f'{seq}: ??')

            # Could be mins or months
            elif len(seq) == 2:
                seq_analysis_list.append(f'{seq}: ??')

            # All other lengths are always months
            elif len(seq) == 3:  # Shortened month name
                seq_analysis_list.append(f'{seq}: Jul')
            elif len(seq) == 4:  # Full month name
                seq_analysis_list.append(f'{seq}: July')
            elif len(seq) == 5:  # Month initial only
                seq_analysis_list.append(f'{seq}: J')
            else:  # Default to full month name for any longer strings
                seq_analysis_list.append(f'{seq}: default - July')

        output_list.append(f'{i}. Sequences found: {seq_analysis_list}')

    else:  # No 'm' sequences are present
        output_list.append(f'{i}. No sequences found')

for statement in output_list:
    print(statement)
