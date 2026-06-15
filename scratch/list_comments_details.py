import re

with open('notebook.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines, 1):
    match = re.search(r'#(Oct\d+|Nov\d+|Feb\d+|NEW HERE|Dec\d+|Jan\d+|March\d+)', line, re.IGNORECASE)
    if match:
        print(f"--- Line {i}: {line.strip()} ---")
        for j in range(i, min(i + 5, len(lines))):
            print(f"  {lines[j].strip()}")
