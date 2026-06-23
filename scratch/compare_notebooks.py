import difflib

with open('notebook.txt', 'r', encoding='utf-8') as f:
    old_lines = f.readlines()

with open('new_notebook.txt', 'r', encoding='utf-8') as f:
    new_lines = f.readlines()

# Run a diff and write it to a file for review
diff = list(difflib.unified_diff(
    old_lines,
    new_lines,
    fromfile='notebook.txt',
    tofile='new_notebook.txt',
    lineterm=''
))

with open('scratch/notebook_diff.txt', 'w', encoding='utf-8') as outf:
    outf.write('\n'.join(diff))

print(f"Diff lines: {len(diff)}")
# Print the first 100 lines of the diff (or check if it's small)
for line in diff[:100]:
    print(line)
