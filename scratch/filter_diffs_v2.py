import re

with open('scratch/diff_report.txt', 'r', encoding='utf-8') as f:
    content = f.read()

# Find all blocks starting with "Diff for"
blocks = re.split(r'===+[\r\n]+', content)

print("Non-identical functions (actual logic diffs):")
for block in blocks:
    block = block.strip()
    if not block:
        continue
    first_line = block.split('\n')[0]
    if "identical" in first_line:
        continue
    print(f"  - {first_line}")
    # Print the line numbers changed
    lines = block.split('\n')
    adds = sum(1 for l in lines if l.startswith('+ ') and not l.startswith('+++ '))
    dels = sum(1 for l in lines if l.startswith('- ') and not l.startswith('--- '))
    print(f"    (dels: {dels}, adds: {adds})")
