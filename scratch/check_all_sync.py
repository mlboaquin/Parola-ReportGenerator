import re
import difflib

with open('main.py', 'r', encoding='utf-8') as f:
    main_code = f.read()

with open('new_notebook.txt', 'r', encoding='utf-8') as f:
    notebook_code = f.read()

def normalize_code(code_str):
    # Remove self.
    code_str = re.sub(r'\bself\.', '', code_str)
    # Remove docstrings
    code_str = re.sub(r'\"\"\"(.*?)\"\"\"', '', code_str, flags=re.DOTALL)
    code_str = re.sub(r"\'\'\'(.*?)\'\'\'", '', code_str, flags=re.DOTALL)
    # Remove comments
    code_str = re.sub(r'#.*', '', code_str)
    # Normalize whitespaces
    lines = [line.strip() for line in code_str.splitlines() if line.strip()]
    return '\n'.join(lines)

def extract_methods_main(code_str):
    # Regex to find def inside class
    matches = re.finditer(r'^\s*def\s+([a-zA-Z0-9_]+)\s*\((.*?)\):', code_str, re.MULTILINE)
    methods = {}
    for match in matches:
        name = match.group(1)
        params = match.group(2)
        start = match.start()
        # Find indent level
        line_start = code_str.rfind('\n', 0, start) + 1
        def_line = code_str[line_start:code_str.find('\n', start)]
        indent = len(def_line) - len(def_line.lstrip())
        
        # Get body lines
        body_lines = [code_str[start:code_str.find('\n', start)]]
        next_lines = code_str[code_str.find('\n', start) + 1:].split('\n')
        for line in next_lines:
            if not line.strip():
                continue
            line_indent = len(line) - len(line.lstrip())
            if line_indent <= indent:
                break
            body_lines.append(line)
        methods[name] = '\n'.join(body_lines)
    return methods

main_methods = extract_methods_main(main_code)
notebook_methods = extract_methods_main(notebook_code)

print(f"main_methods count: {len(main_methods)}")
print(f"notebook_methods count: {len(notebook_methods)}")

different = []
identical = []
for name, n_body in notebook_methods.items():
    if name in main_methods:
        m_body = main_methods[name]
        # Clean both for logical comparison
        m_clean = normalize_code(m_body)
        n_clean = normalize_code(n_body)
        
        # Remove parameters differences (specifically 'self')
        m_clean = re.sub(r'^def\s+\w+\s*\(self,\s*', 'def ' + name + '(', m_clean)
        m_clean = re.sub(r'^def\s+\w+\s*\(self\)', 'def ' + name + '()', m_clean)
        n_clean = re.sub(r'^def\s+\w+\s*\(', 'def ' + name + '(', n_clean)
        
        if m_clean != n_clean:
            different.append(name)
        else:
            identical.append(name)
    else:
        print(f"Method in notebook but not main: {name}")

print(f"\nIdentical: {len(identical)}")
print(f"Different: {len(different)} -> {different}")

# Dump actual diffs for the different ones to check
with open('scratch/mismatch_report.txt', 'w', encoding='utf-8') as f:
    for name in different:
        f.write(f"=== METHOD: {name} ===\n")
        diff = list(difflib.unified_diff(
            main_methods[name].splitlines(),
            notebook_methods[name].splitlines(),
            fromfile='main.py',
            tofile='new_notebook.txt',
            lineterm=''
        ))
        f.write('\n'.join(diff) + "\n\n")
