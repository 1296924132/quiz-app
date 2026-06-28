---
name: docx-template
description: "Word 模板大师 — 参考老师发的 .docx 模板，用 python-docx / docxtpl 快速生成填写好的作业/申报书/报告。支持中文排版、表格填充、批量生成。"
---

# Word 模板大师（docx-template）

## 适用场景

老师发了一个 `.docx` 模板 → 你需要填入自己的内容 → 交作业。

**不适合**：从零排版创作（用 `/docx` skill）。

## 核心原则

> **不要手改模板，要写脚本填模板。**
>
> 老师下次换模板，你改几行数据就能复用。

---

## 🚀 快速上手

### 安装依赖

```bash
pip install python-docx docxtpl
```

### 方法一：`python-docx`（通用方案，适合复杂操作）

```python
from docx import Document

# 1. 打开模板
doc = Document('老师发的模板.docx')

# 2. 遍历段落，替换占位符
for p in doc.paragraphs:
    for run in p.runs:
        if '{{姓名}}' in run.text:
            run.text = run.text.replace('{{姓名}}', '姜沅其')
        if '{{学号}}' in run.text:
            run.text = run.text.replace('{{学号}}', '2024XXXX')

# 3. 遍历表格（很多模板用表格排版）
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if '{{专业}}' in run.text:
                        run.text = run.text.replace('{{专业}}', '计算机科学与技术')

# 4. 保存
doc.save('已填写-作业.docx')
```

### 方法二：`docxtpl`（推荐，Jinja2 模板语法，更优雅）

先在模板里用 `{{ 变量名 }}` 标记占位符，然后：

```python
from docxtpl import DocxTemplate

doc = DocxTemplate('老师发的模板.docx')

context = {
    '姓名': '姜沅其',
    '学号': '2024XXXX',
    '学院': '计算机学院',
    '专业': '计算机科学与技术',
    '日期': '2026年6月24日',
}

doc.render(context)
doc.save('已填写-作业.docx')
```

**优势**：一行 `doc.render(context)` 搞定所有替换，不用手写循环。

---

## 📖 模板分析（第一步一定是这个）

拿到老师发的模板，先分析结构：

```bash
python -c "
from docx import Document
doc = Document('模板.docx')

print('=== 段落 ===')
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f'[{i}] 样式={p.style.name} | {p.text[:60]}')

print('\n=== 表格 ===')
for ti, table in enumerate(doc.tables):
    print(f'\n--- 表格 {ti} ({len(table.rows)}行 x {len(table.columns)}列) ---')
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            text = cell.text.strip()[:40]
            if text:
                print(f'  [{ri},{ci}] {text}')
"
```

> 这步帮我分析模板，我再写填充脚本。

---

## 🎯 常用模式

### 1. 替换文本框 / 表格中的中文占位符

| 占位符风格 | 示例 |
|---|---|
| 双花括号（推荐） | `{{姓名}}` `{{学号}}` |
| 下划线填空 | `________` `______` |
| 标签式 | `[姓名]` `<姓名>` |
| 原样文字 | 直接替换"请填写姓名" |

脚本通用写法：

```python
import re
from docx import Document
from docx.oxml.ns import qn

doc = Document('模板.docx')

replacements = {
    '{{姓名}}': '姜沅其',
    '{{学号}}': '2024XXXX',
    '{{学院}}': '计算机学院',
}

def replace_in_paragraph(p, mapping):
    """替换段落中的占位符"""
    for run in p.runs:
        for old, new in mapping.items():
            if old in run.text:
                run.text = run.text.replace(old, new)

def replace_in_table(table, mapping):
    """递归替换表格中的占位符"""
    for row in table.rows:
        for cell in row.cells:
            # 替换单元格内段落
            for p in cell.paragraphs:
                replace_in_paragraph(p, mapping)
            # 递归处理嵌套表格
            for t in cell.tables:
                replace_in_table(t, mapping)

# 替换所有段落
for p in doc.paragraphs:
    replace_in_paragraph(p, replacements)

# 替换所有表格
for table in doc.tables:
    replace_in_table(table, replacements)

doc.save('已填写.docx')
```

### 2. 设置中文字体（非常重要的坑）

`python-docx` 默认不设中文字体，Word 打开可能显示为西文字体：

```python
from docx.oxml.ns import qn

def set_cn_font(run, font_name='宋体'):
    """设置中文字体（同时设西文和中文字体名）"""
    run.font.name = font_name
    run.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

# 用法
run = p.add_run('你好')
set_cn_font(run, '宋体')
run.font.size = Pt(12)

# 批量设置段落中所有 run 的字体
def set_para_font(para, font_name='宋体', size=Pt(12)):
    for run in para.runs:
        set_cn_font(run, font_name)
        run.font.size = size
```

### 3. 处理表格（模板常用）

```python
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_cell(cell, text, font_name='宋体', size=Pt(10), bold=False,
             align=WD_ALIGN_PARAGRAPH.CENTER):
    """设置单元格内容和格式"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    set_cn_font(run, font_name)
    run.font.size = size
    run.font.bold = bold
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def set_cell_left(cell, text, font_name='宋体', size=Pt(10), bold=False):
    set_cell(cell, text, font_name, size, bold, WD_ALIGN_PARAGRAPH.LEFT)
```

### 4. 表格中新增行（当模板行不够时）

```python
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

doc = Document('模板.docx')
table = doc.tables[0]

# 复制第1行的格式，在后面追加
ref_row = table.rows[1]
new_row = table.add_row()

# 把新行的各单元格格式设为与参考行一致
for i, cell in enumerate(new_row.cells):
    # 复制宽度
    cell.width = ref_row.cells[i].width
    # 设置内容
    set_cell(cell, f'新数据{i}')
```

### 5. `docxtpl` 进阶：循环和条件

在模板中用 `{{ 变量 }}` 的同时，`docxtpl` 还支持：

**在模板中写循环**（比如成员列表）：
```
{{ 团队成员信息 }}
{% for m in members %}
{{ m.姓名 }}    {{ m.学号 }}    {{ m.专业 }}
{% endfor %}
```

**Python 脚本**：
```python
from docxtpl import DocxTemplate

doc = DocxTemplate('模板.docx')
doc.render({
    '姓名': '姜沅其',
    'members': [
        {'姓名': '张三', '学号': '01', '专业': '计算机'},
        {'姓名': '李四', '学号': '02', '专业': '软件工程'},
    ]
})
doc.save('已填写.docx')
```

> ⚠️ `docxtpl` 的循环/条件需要用 RichText 或放在表格特定单元格中。复杂表格建议用纯 `python-docx`。

---

## 🧰 实用脚本工具箱

### 脚本一：分析模板结构

```python
"""analyze_template.py — 分析模板结构，打印段落和表格"""
import sys
from docx import Document

path = sys.argv[1] if len(sys.argv) > 1 else '模板.docx'
doc = Document(path)

print(f'📄 文件: {path}')
print(f'📝 段落数: {len(doc.paragraphs)}')
print(f'📊 表格数: {len(doc.tables)}')
print(f'📐 页面大小: {doc.sections[0].page_width}, {doc.sections[0].page_height}')
print(f'📏 页边距: T={doc.sections[0].top_margin} B={doc.sections[0].bottom_margin} '
      f'L={doc.sections[0].left_margin} R={doc.sections[0].right_margin}')
print()

print('=== 段落内容 ===')
for i, p in enumerate(doc.paragraphs):
    t = p.text.strip()
    if t:
        print(f'P[{i:3d}] [{p.style.name:15s}] {t[:80]}')

print('\n=== 表格内容 ===')
for ti, table in enumerate(doc.tables):
    print(f'\n--- Table {ti} ({len(table.rows)}r × {len(table.columns)}c) ---')
    for ri, row in enumerate(table.rows):
        texts = []
        for ci, cell in enumerate(row.cells):
            t = cell.text.strip()[:25]
            texts.append(t)
        print(f'  R[{ri:2d}] ' + ' | '.join(texts))
```

### 脚本二：批量替换填空线

很多模板用 `________` 做填空线，批量替换成实际内容：

```python
from docx import Document

doc = Document('模板.docx')

# 填空映射：行号 → 内容（行号用 analyze_template.py 查看）
line_fills = {
    5: '姜沅其',        # 姓名
    8: '计算机学院',     # 学院
    12: '2024XXXX',      # 学号
}

filled_paras = 0
for i, p in enumerate(doc.paragraphs):
    if i in line_fills and '____' in p.text:
        for run in p.runs:
            if '____' in run.text:
                run.text = run.text.replace('____', line_fills[i])
                filled_paras += 1

print(f'✅ 填充了 {filled_paras} 个段落')
doc.save('已填写.docx')
```

---

## 🧪 QA 检查（每次生成后必做）

```bash
# 1. 检查有没有未替换的占位符
python -c "
from docx import Document
import re
doc = Document('输出.docx')
unfilled = []
for i, p in enumerate(doc.paragraphs):
    for run in p.runs:
        if re.search(r'\{|\%\{|____|请填写', run.text):
            unfilled.append((i, run.text[:50]))
if unfilled:
    print('❌ 有未填充的占位符:')
    for i, t in unfilled:
        print(f'  P[{i}] {t}')
else:
    print('✅ 所有占位符已填充')
"
```

---

## 📋 工作流程总结

```
老师发模板.docx
      ↓
[1] python analyze_template.py 模板.docx
    → 看懂模板结构（哪些段落、哪些表格）
      ↓
[2] 让 Claude 帮我写填充脚本
    → 告诉它：模板结构 + 我要填什么内容
      ↓
[3] python fill_template.py
    → 生成 已填写.docx
      ↓
[4] QA检查 → 确认无遗漏
      ↓
[5] 交作业 ✅
```

---

## 🔧 常见问题

| 问题 | 解决 |
|---|---|
| 中文显示成方框 | 没设 `w:eastAsia` 字体，用 `set_cn_font()` 修复 |
| 表格里替换不生效 | 表格的 text 在 `cell.paragraphs` 里，不是 `doc.paragraphs` |
| 替换后有占位符残留 | run 被拆成了多段，用 `replace_in_paragraph()` 覆盖所有 run |
| 合并单元格后数据错位 | 合并后的 cell 对象变了，**先合并再填充** |
| docxtpl 不支持复杂表格 | 改用纯 `python-docx`，手动控制每个单元格 |
| 图片/水印未复制 | `python-docx` 打开已有模板默认保留图片，除非另存为新文档 |
