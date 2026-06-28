---
name: python-script
description: Python 脚本开发助手 — 快速生成脚本骨架、统一风格、一键调试
---

# Python 脚本开发助手

## 适用场景

- 写一个新的 Python 工具/脚本（数据处理、文件操作、爬虫、自动化）
- 需要统一项目内 Python 代码风格和结构
- 已有脚本需要调试、重构或扩展

## 核心原则

1. **先看项目现有风格** — 检查项目内其他 Python 脚本的写法（命令行参数、错误处理、输出格式）
2. **骨架先行** — 先输出完整的脚本骨架让用户确认，再填充业务逻辑
3. **可复现** — 所有依赖写入 `requirements.txt`，用 `if __name__ == '__main__':` 入口

---

## 通用脚本骨架模板

```python
#!/usr/bin/env python3
"""脚本用途一句话说明

详细说明：输入什么、输出什么、依赖什么。
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Optional


def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='脚本说明')
    parser.add_argument('input', help='输入文件路径')
    parser.add_argument('-o', '--output', default='output.json', help='输出文件路径')
    parser.add_argument('-v', '--verbose', action='store_true', help='详细输出')
    return parser.parse_args()


def main() -> None:
    """主函数"""
    args = parse_args()

    # 检查输入是否存在
    input_path = Path(args.input)
    if not input_path.exists():
        print(f'❌ 输入文件不存在: {input_path}')
        sys.exit(1)

    # 确保输出目录存在
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # TODO: 核心逻辑

    print(f'✅ 完成，输出到: {output_path}')


if __name__ == '__main__':
    main()
```

---

## 项目风格指南（基于现有项目分析）

项目内 Python 脚本的通用习惯：

| 特征 | 规范 |
|------|------|
| 错误处理 | 用 `print(f'❌ ...')` 输出错误，`sys.exit(1)` 退出 |
| 成功提示 | 用 `print(f'✅ ...')` 输出 |
| 进度提示 | 用 `print(f'⏳ ...')` 或 `print(f'📝 ...')` |
| 路径处理 | 用 `pathlib.Path`，不用 `os.path` |
| 参数解析 | 用 `argparse` |
| 入口 | `if __name__ == '__main__': main()` |
| 字符编码 | 读写文件显式指定 `encoding='utf-8'` |
| 依赖管理 | 用 `requirements.txt` |

---

## 常用脚本模式

### 读取 JSON / 处理数据

```python
data_path = Path('data.json')
data = json.loads(data_path.read_text(encoding='utf-8'))
# 处理...
Path('output.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
```

### 批量处理目录

```python
input_dir = Path('input/')
for f in sorted(input_dir.glob('*.txt')):
    print(f'📝 处理: {f.name}')
    # 处理逻辑...
```

### 生成 Word / Excel

```python
# Word → from docx import Document
# Excel → from openpyxl import Workbook
```

---

## 调试与验证

每次生成脚本后运行验证：

```bash
# 1. 语法检查
python -m py_compile 脚本.py

# 2. 试运行
python 脚本.py --help

# 3. 先 dry-run 确认输出预览再实际执行
```

## 工作流

```
用户需求
  ↓
[1] 检查项目现有脚本风格
  ↓
[2] 生成脚本骨架 → 用户确认
  ↓
[3] 填充业务逻辑
  ↓
[4] 语法检查 + 试运行
  ↓
[5] 正式执行 → 输出结果
```
