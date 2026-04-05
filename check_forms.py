
import re

with open(r'c:\Users\t_kou\Kakeibo_Final_v3\app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

form_blocks = []
current_form = None

for i, line in enumerate(lines):
    if 'with st.form(' in line:
        indent = len(line) - len(line.lstrip())
        current_form = {'start': i + 1, 'indent': indent, 'has_submit': False, 'name': line.strip()}
        form_blocks.append(current_form)
    elif current_form:
        line_indent = len(line) - len(line.lstrip())
        if line.strip() and line_indent <= current_form['indent'] and i + 1 > current_form['start']:
            # block ended
            current_form = None
        elif 'st.form_submit_button(' in line:
            current_form['has_submit'] = True

for form in form_blocks:
    if not form['has_submit']:
        print(f"Form at line {form['start']} ({form['name']}) is missing a submit button.")
    else:
        print(f"Form at line {form['start']} ({form['name']}) has a submit button.")
