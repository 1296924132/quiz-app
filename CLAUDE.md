# TREA — 项目根目录

## 📁 完整目录结构

```
TRAE文件/                          ← Git 根目录 / 工作目录
│
├── 申报书生成/                     # 📄 暑期社会实践申报书
│   ├── create_filled_form.py       #   主脚本：生成申报书
│   ├── generate_competition_form.py
│   ├── requirements.txt
│   └── run.sh                     #   一键运行
│
├── 刷题神器/                       # 🎯 公文写作刷题前端应用（纯前端 HTML）
│   ├── 公文写作刷题神器（离线版）.html  #   主应用
│   ├── 刷题神器.html                    #   轻量加载版（需同目录JSON）
│   ├── index.html
│   └── debug/                      #   数据处理脚本（归档）
│
├── 反恐实训作业/                    # 👮 反恐怖现场处置实训作业
│   ├── fill_report_v3.py           #   生成脚本 V3
│   ├── polish_report.py            #   润色脚本
│   └── 加分单模版.docx              #   Word 模板
│
├── 暑假实习意向分析/                 # 📊 暑假实习意向汇总分析
│   ├── generate_report.py          #   生成分析报告
│   ├── generate_summary.py         #   生成汇总表
│   ├── summary_data.json
│   ├── 暑假实习意向分析报告.docx
│   ├── 暑假实习意向汇总表.xlsx
│   └── 暑假实习意向表/              #   各班原始数据
│
├── PPT/                            # 📽️ PPT 生成
│   └── create_ppt.py
│
├── 资料/                           # 📚 原始素材
│   └── 学习通 基础写作题库.docx
│
├── .claude/                        # 🤖 Claude 配置
│   ├── skills/                     #   项目级 skill
│   ├── scripts/                    #   辅助工具脚本
│   ├── hooks/                      #   生命周期钩子
│   ├── settings.json               #   项目配置
│   └── memory/                     #   记忆文件（已 gitignore）
│
├── CLAUDE.md                       # 本文件
├── README.md
└── .gitignore
```

---

## 🚀 快速启动

### 申报书生成

```bash
cd 申报书生成
pip install -r requirements.txt && python create_filled_form.py
```

输出：`附件3：2026年暑期社会实践活动院级项目申报书（已填写）.docx`

> `.claude/skills/generate-form` 中已配置同名 skill，可用 `/generate-form` 调用。

### 刷题神器

直接在浏览器打开 `刷题神器/公文写作刷题神器（离线版）.html`。

### PPT 生成

```bash
python PPT/create_ppt.py
```

### 反恐实训作业

```bash
cd 反恐实训作业
pip install -r requirements.txt
python fill_report_v3.py
```

### 暑假实习意向分析

```bash
cd 暑假实习意向分析
pip install -r requirements.txt
python generate_summary.py     # 生成汇总表
python generate_report.py      # 生成分析报告
```

---

## 🛠 通用工具

| Skill | 用途 | 调用 |
|-------|------|------|
| 申报书生成 | 生成社会实践申报书 | `/generate-form` |
| Word 模板填充 | 分析模板 → 写填充脚本 | `/docx-template` |
| PPT 制作 | 创建/编辑/美化 PPT | `/ppt-master` |
| 自治开发循环 | Ralph 循环反复执行 | `/ralph-loop` |

## ⚠️ 工作目录说明

**工作目录必须设置为 `D:\TRAE文件`**（Git 根目录），**不要**设为 `.claude/scripts/` 或其他子目录，否则会导致配置嵌套。
