"""Fix indentation in process_customers method"""

with open(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_bulk_processor.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the for loop in process_customers (around line 535)
# Everything from after "try:" at line 569 until the except blocks needs to be dedented by 4 spaces

output_lines = []
in_for_body = False
for_start_line = None

for i, line in enumerate(lines):
    line_num = i + 1
    
    # Detect the for loop
    if 'for index, row in self.df.iterrows():' in line and not in_for_body:
        in_for_body = True
        for_start_line = line_num
        output_lines.append(line)
        continue
    
    # Detect end of for loop (dedent back to "return self.results")
    if in_for_body and line.strip() == 'return self.results':
        in_for_body = False
        output_lines.append(line)
        continue
    
    # If we're in the for body, handle the "try:" and everything after it
    if in_for_body and line_num >= 569:
        # Dedent by 4 spaces if line starts with 16 spaces
        if line.startswith('                '):  # 16 spaces
            output_lines.append(line[4:])  # Remove 4 spaces
        else:
            output_lines.append(line)
    else:
        output_lines.append(line)

# Write back
with open(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_bulk_processor.py', 'w', encoding='utf-8') as f:
    f.writelines(output_lines)

print('Fixed indentation!')
