import re
import difflib

with open('main.py', 'r', encoding='utf-8') as f:
    main_content = f.read()

with open('new_notebook.txt', 'r', encoding='utf-8') as f:
    notebook_content = f.read()

def get_function_body(content, name):
    match = re.search(r'def\s+' + name + r'\b', content)
    if not match:
        return None
    start_idx = match.start()
    line_start = content.rfind('\n', 0, start_idx) + 1
    def_line = content[line_start:content.find('\n', start_idx)]
    indent = len(def_line) - len(def_line.lstrip())
    
    lines = content[start_idx:].split('\n')
    body_lines = [lines[0]]
    for line in lines[1:]:
        if not line.strip():
            body_lines.append(line)
            continue
        line_indent = len(line) - len(line.lstrip())
        if line_indent <= indent:
            break
        body_lines.append(line)
    return '\n'.join(body_lines)

name = 'ensure_orr_header_and_spacing'
m_body = get_function_body(main_content, name)
n_body = get_function_body(notebook_content, name)

# Save diff to a text file for complete inspection
diff = list(difflib.unified_diff(
    m_body.splitlines(),
    n_body.splitlines(),
    fromfile='main.py',
    tofile='new_notebook.txt',
    lineterm=''
))

with open('scratch/orr_header_spacing_diff.txt', 'w', encoding='utf-8') as outf:
    outf.write('\n'.join(diff))

print(f"Diff written to scratch/orr_header_spacing_diff.txt. Diff length: {len(diff)}")
