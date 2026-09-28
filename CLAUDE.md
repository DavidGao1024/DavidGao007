# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是高佳顺的个人简历项目，包含三种格式的简历文件，内容保持同步。

## 文件结构

| 文件 | 用途 |
|------|------|
| `高佳顺简历_优化版.md` | Markdown 源文件，主要编辑入口 |
| `高佳顺简历.docx` | Word 格式，用于发送给HR |
| `index.html` | 静态网页版，用于线上展示 |
| `create_docx.py` | Python 脚本，用于生成 Word 文档 |

## 常用命令

### 生成 Word 文档
```bash
python create_docx.py
```
需要安装依赖：`pip install python-docx`

### 预览网页版
直接在浏览器中打开 `index.html`

## 内容同步规则

修改简历内容时，需要同步更新三个文件：
1. `高佳顺简历_优化版.md` - Markdown 源文件
2. `index.html` - HTML 网页
3. `create_docx.py` - 修改后重新运行生成 Word

## 简历结构

1. 基本信息（姓名、联系方式、期望薪资）
2. 核心优势（6项）
3. 技术栈（表格形式）
4. AI技能与工具（独立板块）
5. 工作经历（时间线，2009-至今）
6. 代表项目（7个项目）
7. 教育背景
