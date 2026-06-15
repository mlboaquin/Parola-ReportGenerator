import re

def find_comments(filepath):
    comments = []
    with open(filepath, 'r', encoding='utf-8') as f:
        for i, line in enumerate(f, 1):
            match = re.search(r'#(Oct\d+|Nov\d+|Feb\d+|NEW HERE|Dec\d+|Jan\d+|March\d+)', line, re.IGNORECASE)
            if match:
                comments.append((i, match.group(0), line.strip()))
    return comments

notebook_comments = find_comments('notebook.txt')
main_comments = find_comments('main.py')

print(f"Total date/NEW comments in notebook.txt: {len(notebook_comments)}")
print(f"Total date/NEW comments in main.py: {len(main_comments)}")

print("\nSample comments from notebook.txt:")
for idx, tag, line in notebook_comments[:20]:
    print(f"  Line {idx}: {line}")

print("\nSample comments from main.py:")
for idx, tag, line in main_comments[:20]:
    print(f"  Line {idx}: {line}")
