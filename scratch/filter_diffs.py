with open('scratch/diff_report.txt', 'r', encoding='utf-8') as f:
    content = f.read()

sections = content.split('================================================================================\n')
print(f"Total sections: {len(sections)}")

print("\nNon-identical functions (actual logic diffs):")
for sec in sections:
    if not sec.strip():
        continue
    first_line = sec.split('\n')[0]
    if "is identical" not in sec:
        print(f"  - {first_line}")
        # print first few lines of the diff
        lines = sec.split('\n')
        diff_lines = [l for l in lines if l.startswith('+ ') or l.startswith('- ')]
        print(f"    (Diff size: {len(diff_lines)} lines modified)")
