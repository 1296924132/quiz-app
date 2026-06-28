import json, re

with open('tmp_gen_explanations.py', 'r', encoding='utf-8') as f:
    content = f.read()

data_start = content.find('data = [')
result_start = content.find('result = []')
data_section = content[data_start:result_start]

# Try to parse each line that looks like a data entry
lines = data_section.split('\n')
problem_lines = []
ok_lines = 0
for i, line in enumerate(lines):
    stripped = line.strip()
    if not stripped or stripped in ('data = [', ']'):
        continue
    # Remove trailing comma
    test = stripped.rstrip(',')
    try:
        json.loads(test)
        ok_lines += 1
    except json.JSONDecodeError as e:
        problem_lines.append((i, stripped[:200], str(e)))

print(f'OK lines: {ok_lines}')
print(f'Problem lines: {len(problem_lines)}')
print()
for idx, txt, err in problem_lines[:15]:
    print(f'Line {idx}:')
    print(f'  Error: {err}')
    print(f'  Text: {txt[:180]}')
    print()
