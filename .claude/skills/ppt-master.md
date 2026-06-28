---
name: ppt-master
description: PPT Master — PowerPoint 幻灯片大师，创建、编辑、美化 PPT
---

# PPT Master

PowerPoint 幻灯片大师技能。基于 `python-pptx` 和系统 `pptx` skill，提供从创建到美化的全流程支持。

## 快速命令

### 运行项目已有脚本

```bash
python PPT/create_ppt.py
```

### 从零创建 PPT

使用系统 `pptx` skill 的 `pptxgenjs` 方案（推荐模板缺失时）：

```bash
npm install -g pptxgenjs
```

然后生成 `pptxgenjs` 脚本 → `node script.js` 输出 `.pptx`。

### 读取/分析 PPT

```bash
python -m markitdown 文件名.pptx
```

### PPT 转图片做视觉检查

```bash
python /c/Users/12969/.claude/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf 文件名.pptx
pdftoppm -jpeg -r 150 文件名.pdf slide
```

## 设计原则

### 每页必做
- 每页幻灯片必须有视觉元素（图标、图片、形状）— 纯文字页不可取
- 不要重复相同布局 — 各页之间要变化
- 不要用下划线装饰标题 — 它是 AI 生成幻灯片的标志

### 配色
- 选有主题感的配色，不拘泥于蓝色
- 主色占 60-70%，辅色 1-2 种，强调色 1 种
- 深色背景用于封面/结尾页，浅色用于内容页（"三明治"结构）

### 字体
- 标题用 Georgia / Arial Black / Calibri（36-44pt）
- 正文用 Calibri / 微软雅黑（14-16pt）
- 说明文字 10-12pt

### 间距
- 最小边距 0.5"
- 模块之间留 0.3-0.5" 空隙
- 不要填满每一寸空间

## QA 检查（必须）

每次生成后必须做 QA：

1. 内容检查：`python -m markitdown output.pptx`，确认没有丢失内容、占位符文本
2. 视觉检查：转图片后用 subagent 找重叠、溢出、对比度等问题
3. 修复后重新验证 — 至少完成一輪 fix→verify 才能说完成

## 系统级 pptx skill

本 skill 是项目级快捷入口。完整 PPT 能力（编辑模板、高级设计、解包/打包等）由系统级 `pptx` skill 提供，需要时用 `/pptx` 调用。
