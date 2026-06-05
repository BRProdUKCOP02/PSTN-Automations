"""Fix indentation in phoneline_plus_bulk_processor.py"""

# Read the file
with open(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_bulk_processor.py', 'r', encoding='utf-8') as f:
    content = f.read()

# The issue is that everything from line 674 onwards (after "row_num = ...") 
# until before the "finally:" at line 964 needs to be indented by 4 spaces

lines = content.split('\n')

# Find the for loop
for_line_idx = None
for i, line in enumerate(lines):
    if 'for index, row in self.df.iterrows():' in line:
        for_line_idx = i
        break

if for_line_idx is None:
    print("Could not find for loop!")
    exit(1)

print(f"Found for loop at line {for_line_idx + 1}")

# Find the matching finally (at same indentation as the try that contains the for)
try_indent = None
for i in range(for_line_idx - 1, 0, -1):
    if lines[i].strip().startswith('try:'):
        try_indent = len(lines[i]) - len(lines[i].lstrip())
        print(f"Found try at line {i+1} with indent {try_indent}")
        break

finally_line_idx = None
for i in range(for_line_idx + 1, len(lines)):
    if lines[i].strip().startswith('finally:'):
        current_indent = len(lines[i]) - len(lines[i].lstrip())
        if current_indent == try_indent:
            finally_line_idx = i
            print(f"Found finally at line {i+1} with indent {current_indent}")
            break

if finally_line_idx is None:
    print("Could not find finally!")
    exit(1)

# Now indent everything between for_line and finally_line that needs it
for_indent = len(lines[for_line_idx]) - len(lines[for_line_idx].lstrip())
expected_body_indent = for_indent + 4

print(f"For loop indent: {for_indent}, expected body indent: {expected_body_indent}")

output_lines = []
for i, line in enumerate(lines):
    if i <= for_line_idx or i >= finally_line_idx:
        # Outside the for loop, keep as-is
        output_lines.append(line)
    else:
        # Inside the for loop body
        if not line.strip():
            # Keep empty lines empty
            output_lines.append(line)
        else:
            current_indent = len(line) - len(line.lstrip())
            # Calculate how much to adjust
            if current_indent < expected_body_indent:
                # Need more indent
                additional = expected_body_indent - current_indent
                output_lines.append(' ' * additional + line)
            else:
                # Already indented enough (or nested deeper), keep as-is
                output_lines.append(line)

#Write back
with open(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_bulk_processor.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print("Fixed indentation!")
