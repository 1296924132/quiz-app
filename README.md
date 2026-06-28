# TREA

多项目聚合工作区。

## 📁 项目一览

| 子目录 | 说明 | 技术栈 |
|--------|------|--------|
| [申报书生成](申报书生成/) | 暑期社会实践院级项目申报书 | Python + python-docx |
| [刷题神器](刷题神器/) | 公文写作刷题离线前端 | 纯 HTML/CSS/JS |
| [反恐实训作业](反恐实训作业/) | 反恐怖现场处置实训期末作业 | Python + python-docx |
| [暑假实习意向分析](暑假实习意向分析/) | 暑期实习意向汇总分析 | Python + openpyxl + pandas |
| [PPT](PPT/) | 幻灯片生成脚本 | Python + python-pptx |
| [资料](资料/) | 原始素材（学习通题库等） | — |

## 快速启动

```bash
# 申报书生成
cd 申报书生成 && pip install -r requirements.txt && python create_filled_form.py

# 反恐实训作业
cd 反恐实训作业 && pip install python-docx && python fill_report_v3.py

# 暑假实习意向分析
cd 暑假实习意向分析 && pip install openpyxl pandas python-docx
python generate_summary.py && python generate_report.py

# PPT 生成
python PPT/create_ppt.py
```

## 工作目录

**必须设为 `D:\TRAE文件`**（Git 根目录）。
