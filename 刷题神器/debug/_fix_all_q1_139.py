import json, os, glob, re, subprocess, base64, urllib.request, ssl

# ===== 1. 从原始docx提取所有正确答案 =====
files = glob.glob('*.docx'); fname = os.path.abspath(files[0])
with open(fname, 'r', encoding='gb18030') as f: c = f.read()
clean = re.sub(r'<[^>]+>', '\n', c)
clean = clean.replace('&amp;','&').replace('&lt;','<').replace('&gt;','>')
clean = clean.replace('&nbsp;',' ').replace('&quot;','"')
lines = [l.strip() for l in clean.split('\n') if l.strip()]
single_start = None
for i, line in enumerate(lines):
    if line == '单选题': single_start = i + 1; break
section_3 = None
for i, line in enumerate(lines):
    if line == '三、': section_3 = i; break

# 提取所有题目的原始答案+题干
raw = {}
i = single_start
while i < section_3:
    m = re.match(r'^(\d+)\.$', lines[i])
    if m:
        qnum = int(m.group(1)); i += 1
        q_parts = []; answer = ''
        while i < section_3:
            line = lines[i]
            if re.match(r'^\d+\.$', line): break
            if '答案' in line:
                i += 1
                if i < section_3 and re.match(r'^[A-D]+$', lines[i]):
                    answer = lines[i]; i += 1
                else:
                    answer = line.replace('答案：','').replace('答案:','').replace('答案','').strip()
                break
            q_parts.append(line); i += 1
        raw[qnum] = {'question': ' '.join(q_parts).strip(), 'answer': answer}
    else: i += 1

# ===== 2. 从git原始版获取正确解析 =====
result = subprocess.run(['git', 'show', '04f5bb6:公文写作刷题神器（离线版）.html'],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
git_html = result.stdout
gi = git_html.find('const QUESTION_DATA = '); ge = git_html.find(';\n', gi)
git_data = json.loads(git_html[gi + len('const QUESTION_DATA = '):ge])

# 用agent解析
with open('_explanation_single.json', 'r', encoding='utf-8') as f:
    agent_exps = json.load(f)
agent_map = {item['id']: item['explanation'] for item in agent_exps}

# 建立git题干→agent解析的映射
def clean_key(t):
    return t[:20].strip().replace(' ','').replace('（','').replace('）','').replace('(','').replace(')','').replace('，','').replace('。','')

text_to_exp = {}
for q in git_data['singles']:
    k = clean_key(q['question'])
    oid = q['id']
    if oid in agent_map and agent_map[oid]:
        text_to_exp[k] = agent_map[oid]
print(f'文本→解析映射: {len(text_to_exp)} 条')

# ===== 3. 读取当前HTML =====
with open('../公文写作刷题神器（离线版）.html', 'r', encoding='utf-8') as f:
    h = f.read()
vi = h.find('const QUESTION_DATA = '); ve = h.find(';\n', vi)
d = json.loads(h[vi + len('const QUESTION_DATA = '):ve])

# ===== 4. 逐一修正Q1-Q139 =====
fix_ans = []
fix_exp = []

for q in d['singles']:
    if q['id'] > 139: break

    # 修正答案
    if q['id'] in raw:
        raw_ans = raw[q['id']]['answer'].strip().upper()
        cur_ans = q['answer'].strip().upper()
        if raw_ans and cur_ans != raw_ans:
            fix_ans.append((q['id'], cur_ans, raw_ans, q['question'][:30]))
            q['answer'] = raw_ans

    # 修正解析
    k = clean_key(q['question'])
    if k in text_to_exp:
        q['explanation'] = text_to_exp[k]
        fix_exp.append(q['id'])
    else:
        for tk, te in text_to_exp.items():
            if q['question'][:15].replace(' ','') in tk or tk[:15] in q['question'].replace(' ',''):
                q['explanation'] = te
                fix_exp.append(q['id'])
                break

print(f'\n答案修正: {len(fix_ans)} 道')
for qid, cur, raw_a, qt in fix_ans:
    print(f'  Q{qid}: {cur}→{raw_a} | {qt}')

print(f'\n解析修正: {len(fix_exp)} 道')

# ===== 5. 保存并推送 =====
new_str = json.dumps(d, ensure_ascii=False, separators=(',', ':'))
h = h[:vi + len('const QUESTION_DATA = ')] + new_str + h[ve:]
with open('公文写作刷题神器（离线版）.html', 'w', encoding='utf-8') as f:
    f.write(h)

# 推送
proc = subprocess.run(['git', 'credential-manager', 'get'],
                     input='host=github.com\nprotocol=https\n\n',
                     capture_output=True, text=True, encoding='utf-8', errors='replace')
token = None
for line in proc.stdout.split('\n'):
    if line.startswith('password='): token = line[9:].strip(); break

encoded = base64.b64encode(h.encode('utf-8')).decode('ascii')
url = 'https://api.github.com/repos/1296924132/quiz-app/contents/index.html'
headers = {'Authorization': f'Bearer {token}', 'Accept':'application/vnd.github+json','Content-Type':'application/json'}
ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
req = urllib.request.Request(url, headers=headers, method='GET')
sha = json.loads(urllib.request.urlopen(req, context=ctx).read())['sha']
payload = json.dumps({'message': 'fix: 修正Q1-Q139答案和解析', 'content': encoded, 'sha': sha, 'branch': 'master'}).encode('utf-8')
req = urllib.request.Request(url, data=payload, headers=headers, method='PUT')
urllib.request.urlopen(req, context=ctx)
print('\n✅ 已推送！')
