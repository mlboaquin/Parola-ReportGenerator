import re
import difflib

# Let's read both files
with open('main.py', 'r', encoding='utf-8') as f:
    main_content = f.read()

with open('notebook.txt', 'r', encoding='utf-8') as f:
    notebook_content = f.read()

# Let's extract functions and check diffs for ones that we suspect have changed
# We will write a diff report to a text file
report = []

def get_function_body(content, name):
    # Find def name(
    match = re.search(r'def\s+' + name + r'\b', content)
    if not match:
        return None
    start_idx = match.start()
    # Find indentation of the line with def
    line_start = content.rfind('\n', 0, start_idx) + 1
    def_line = content[line_start:content.find('\n', start_idx)]
    indent = len(def_line) - len(def_line.lstrip())
    
    # Extract lines until we see a line with <= indent that is not empty
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

suspects = [
    'unlock_password_protected_docx',
    'extract_mapping_section',
    'extract_criteria_section',
    'remove_section',
    'insert_element_after',
    'get_short_patent_name_with_suffix',
    'get_short_patent_name_v2',
    'format_patent_display',
    'extract_patent_number',
    'parse_claim_numbers',
    'get_all_claim_numbers_from_google',
    'format_claims_as_ranges',
    'format_date',
    'isUSPatent',
    'fetch_abstract',
    'extract_claim_fragments_from_excel',
    'format_number_with_commas',
    'get_claim_from_google_patents',
    'clone_row_after',
    'clear_cell',
    'set_cell_text',
    'add_hyperlink_to_paragraph',
    'simple_replace_section',
    'fix_document_structure',
    'set_header_font_sizes',
    'ensure_orr_header_and_spacing',
    'remove_stray_orr_heading'
]

for name in suspects:
    m_body = get_function_body(main_content, name)
    n_body = get_function_body(notebook_content, name)
    
    if m_body is None and n_body is None:
        report.append(f"Function {name} not found in either file.\n")
    elif m_body is None:
        report.append(f"Function {name} ONLY in notebook.txt:\n{n_body}\n" + "="*80 + "\n")
    elif n_body is None:
        report.append(f"Function {name} ONLY in main.py:\n{m_body}\n" + "="*80 + "\n")
    else:
        # Check if they are different (ignoring self parameter differences for simplicity, or just showing diff)
        # Normalize self. for comparison
        m_norm = re.sub(r'\bself\.', '', m_body)
        n_norm = n_body
        # normalize whitespace
        m_norm_ws = '\n'.join(l.strip() for l in m_norm.split('\n'))
        n_norm_ws = '\n'.join(l.strip() for l in n_norm.split('\n'))
        
        if m_norm_ws != n_norm_ws:
            diff = list(difflib.unified_diff(
                m_body.splitlines(),
                n_body.splitlines(),
                fromfile='main.py: ' + name,
                tofile='notebook.txt: ' + name,
                lineterm=''
            ))
            report.append(f"Diff for {name}:\n" + '\n'.join(diff) + "\n" + "="*80 + "\n")
        else:
            report.append(f"Function {name} is identical (modulo self).\n" + "="*80 + "\n")

with open('scratch/diff_report.txt', 'w', encoding='utf-8') as outf:
    outf.write('\n'.join(report))

print("Diff report written successfully.")
