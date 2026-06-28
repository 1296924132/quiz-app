import json, re

with open('tmp_gen_explanations.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# --- Extract explanations ---
# Format: key: "text",
explanations = {}
in_explanations = False
for line in lines:
    stripped = line.strip()
    if stripped.startswith('explanations = {'):
        in_explanations = True
        continue
    if in_explanations:
        if stripped == '}' or stripped.startswith('#'):
            continue
        if stripped.startswith('}') or 'data' in stripped:
            break
        # Match: N: "text",
        m = re.match(r'^(\d+):\s*"(.+)"\s*,?\s*$', stripped)
        if m:
            key = int(m.group(1))
            val = m.group(2)
            explanations[key] = val
        else:
            # Try without trailing quote (multi-line?)
            m2 = re.match(r'^(\d+):\s*"(.+)', stripped)
            if m2:
                key = int(m2.group(1))
                val = m2.group(2)
                explanations[key] = val

print(f"Extracted {len(explanations)} explanations")

# --- Extract data list ---
# Each entry is on one line: {"id":N,"question":"...","answer":"X","options":{...}},
# But questions may contain unescaped " which breaks JSON parsing.
# We need to handle this manually.

def extract_json_objects(text):
    """Extract JSON objects from text, handling unescaped quotes in string values."""
    objects = []
    i = 0
    while i < len(text):
        # Find next {
        brace_start = text.find('{', i)
        if brace_start == -1:
            break

        # Find matching }
        depth = 0
        in_string = False
        escape = False
        j = brace_start
        while j < len(text):
            ch = text[j]
            if escape:
                escape = False
                j += 1
                continue
            if ch == '\\' and in_string:
                escape = True
                j += 1
                continue
            if ch == '"':
                in_string = not in_string
                j += 1
                continue
            if not in_string:
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        # Found complete object
                        obj_str = text[brace_start:j+1]
                        # Now parse this JSON with a lenient approach
                        try:
                            obj = json.loads(obj_str)
                            objects.append(obj)
                        except json.JSONDecodeError:
                            # Try to fix embedded quotes
                            fixed = fix_embedded_quotes(obj_str)
                            try:
                                obj = json.loads(fixed)
                                objects.append(obj)
                            except json.JSONDecodeError as e:
                                print(f"Cannot parse: {obj_str[:200]}")
                                print(f"  Error: {e}")
                        break
            j += 1
        i = j + 1 if depth == 0 else brace_start + 1
    return objects

def fix_embedded_quotes(s):
    """Fix unescaped quotes within string values."""
    result = []
    in_string = False
    escape = False
    for ch in s:
        if escape:
            result.append(ch)
            escape = False
            continue
        if ch == '\\':
            result.append(ch)
            escape = True
            continue
        if ch == '"':
            if not in_string:
                in_string = True
                result.append(ch)
            else:
                # Could be end of string or embedded quote
                # For now just keep it and let json.loads decide
                result.append(ch)
                in_string = False
        else:
            result.append(ch)
    return ''.join(result)

data_section = ''.join(lines)
data_start = data_section.find('data = [')
data_end = data_section.find('result = []')
if data_end == -1:
    data_end = data_section.rfind(']')
    if data_end > data_start:
        data_end += 1

raw_data = data_section[data_start:data_end]
data = extract_json_objects(raw_data)
print(f"Extracted {len(data)} questions")

# --- Build output ---
result = []
for item in data:
    qid = item['id']
    if qid in explanations:
        result.append({"id": qid, "explanation": explanations[qid]})
    else:
        print(f"WARNING: No explanation for ID {qid}")

# Write output
with open('_explanation_single.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"Generated _explanation_single.json with {len(result)} entries")
