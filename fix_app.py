with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find and remove the debug block (lines 2950-2971 approximately)
# Look for the pattern: try: st.info("Step 1: Parsing...
new_lines = []
skip_until_step1 = False
found_debug_block = False

i = 0
while i < len(lines):
    line = lines[i]
    
    # Check if this is the start of the debug block
    if 'try:' in line and i + 1 < len(lines) and 'Step 1: Parsing monthly sales data' in lines[i + 1]:
        found_debug_block = True
        # Skip lines until we find "# STEP 1: Parse monthly sales"
        while i < len(lines):
            if '# STEP 1: Parse monthly sales data' in lines[i] or '# ============' in lines[i]:
                # Keep this line and continue normally
                break
            i += 1
        continue
    
    new_lines.append(line)
    i += 1

with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

if found_debug_block:
    print('Debug block removed successfully')
else:
    print('Debug block not found')

print(f'Total lines: {len(new_lines)}')
