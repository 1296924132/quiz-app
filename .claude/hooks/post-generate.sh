#!/bin/bash
# 生成 Word 文件后提示（输出在 申报书生成/ 目录下）
docx_file=$(ls -t "申报书生成/附件3："*.docx 2>/dev/null | head -1)
if [ -n "$docx_file" ]; then
  echo "已生成：$docx_file"
  # 取消注释下一行可在生成后自动打开文件
  # start "$docx_file"
fi
