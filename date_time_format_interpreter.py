import re

dt_formats = ['myhm',
              'yyymms mmm',
              'ddmmyy hhmmmss',
              'yyymm',
              'yyyymm - s mmm',
              'yyhhmm - s mmm ddmm',
              'ddhhyyy',
              'mmmmdddd',
              'hdmsm',
              'yys mm  smm',
              'mmyyh mm  smm m   - s']

output_list = []

for i in range(len(dt_formats)):
    dt_format = dt_formats[i].lower()  # Excel is case insensitive
    if dt_format.count('m') > 0:
        m_sequences = list(re.finditer(r'm+', dt_format))

        seq_analysis_list = []
        found_minute = False  # Track if a minute has been found

        for seq in m_sequences:
            seq_string = seq.group()
            seq_start = seq.start()
            seq_end = seq.end()

            # Could be minutes or months
            if len(seq_string) == 1 or len(seq_string) == 2:
                if not found_minute:
                    # Check for 'h' or 's' preceding the 'm's
                    is_minute = False
                    idx = seq_start - 1

                    while idx >= 0:
                        char = dt_format[idx]
                        if char.isalnum():
                            if char in ('h', 's'):
                                is_minute = True
                            break
                        idx -= 1

                    if is_minute:
                        seq_analysis_list.append(f'{seq_string}: Minute')
                        found_minute = True
                    else:
                        seq_analysis_list.append(f'{seq_string}: Month')
                else:
                    # Default to months unless 's' is found after the 'm' sequence
                    is_minute = False
                    idx = seq_end
                    while idx < len(dt_format):
                        char = dt_format[idx]
                        if char.isalnum():
                            if char == 's':
                                is_minute = True
                            break
                        idx += 1

                    if is_minute:
                        seq_analysis_list.append(f'{seq_string}: Minute')
                        found_minute = True
                    else:
                        seq_analysis_list.append(f'{seq_string}: Month')
                        
            else:
                # All other lengths are always months
                if len(seq_string) == 3:
                    seq_analysis_list.append(f'{seq_string}: Jul')
                elif len(seq_string) == 4:
                    seq_analysis_list.append(f'{seq_string}: July')
                elif len(seq_string) == 5:
                    seq_analysis_list.append(f'{seq_string}: J')
                else:
                    seq_analysis_list.append(f'{seq_string}: default - July')

        output_list.append(f"{i}. Sequences found: {seq_analysis_list}")

    else:
        output_list.append(f"{i}. No 'm' sequences found")

for statement in output_list:
    print(statement)
