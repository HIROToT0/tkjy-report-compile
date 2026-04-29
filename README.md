# 太科检测报告编制系统自动化

> 操作太科检测报告编制系统（comp1.taiketest.com），配合飞书表格实现混凝土试块抗压报告的批量编制、上传与状态回填。

[![Python 3](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-自动化-orange.svg)](https://playwright.dev/)

## 功能特性

- ✅ 从飞书表格读取报告编号清单
- ✅ 自动判断报告状态（报告待编制 / 检验完成 / 报告已上传）
- ✅ 批量编辑样品达到值，支持多组样品并行处理
- ✅ 异常样品自动识别（达到值 < 100 或显示"无效"）
- ✅ 报告自动生成与上传
- ✅ 处理结果回填飞书表格（报告提交状态 / 异常情况说明）

## 系统架构

```
飞书表格（任务清单）
    │
    ▼ 读取报告编号
┌──────────────────────────────────┐
│   太科检测报告编制系统            │
│   comp1.taiketest.com/Report/    │
│                                  │
│  ┌─────────────────────────────┐ │
│  │ 1. 判断报告状态              │ │
│  │ 2. 编辑样品达到值（多组）    │ │
│  │ 3. 异常判定 (<100 → 取消)    │ │
│  │ 4. 生成报告 → 上传报告       │ │
│  └─────────────────────────────┘ │
└──────────────────────────────────┘
    │
    ▼ 回填状态
飞书表格（报告提交状态 / 异常情况说明）
```

## 快速开始

### 环境要求

- Python 3.8+
- Playwright (`pip install playwright && playwright install chromium`)
- Chrome / Chromium 浏览器

### 安装

```bash
# 克隆仓库
git clone https://github.com/HIROToT0/tkjy-report-compile.git
cd tkjy-report-compile

# 安装依赖
pip install playwright
playwright install chromium
```

### 配置凭证

编辑脚本中的凭证配置：

```python
# 太科检测报告系统
USERNAME = "hewei"
PASSWORD = "88833223534qwer32abc@thinks"
COOKIE_FILE = "~/.hermes/report_cookies_comp1.txt"

# 飞书应用
FEISHU_APP_ID = "cli_a95212122abddcee"
FEISHU_APP_SECRET = "qESYjM4BnzyYS7sEPlyCvdMuKOGdypZG"
FEISHU_SPREADSHEET_TOKEN = "NU7wsFP2chQq0htIjUCcHKqVn0d"
FEISHU_SHEET_ID = "ac04d9"
```

### 运行

```bash
python tkjy_report_compile.py
```

## 核心流程

### 判定规则

| 条件 | 操作 |
|------|------|
| 达到值 ≥ 100 | 点击"确定"，正常提交 |
| 达到值 < 100 | 点击"取消"，标记"异常" |
| 显示"无效" | 点击"取消"，标记"异常" |
| 多组样品任一异常 | 整体标记"异常"，跳过剩余组 |

### 报告状态处理

| 状态 | 操作 |
|------|------|
| 报告待编制 | 继续处理 |
| 检验完成 | 继续处理 |
| 报告已上传 | 跳过，记录当前状态 |

## 文件结构

```
tkjy-report-compile/
├── README.md          # 本文件
├── SKILL.md           # 完整操作手册（含避坑指南）
└── scripts/           # 辅助脚本目录
```

## 关联系统

| 系统 | 地址 |
|------|------|
| 太科检测报告系统 | http://comp1.taiketest.com/Report/ |
| 飞书表格（任务清单） | https://ccnlg9zq6b6x.feishu.cn/wiki/WTMWwhfcxiKjvkkLKhFc3haAnOf |

## 常见问题

**Q: 页面停留在登录页，无法自动登录？**
A: 删除 `~/.hermes/report_cookies_comp1.txt`，重新运行脚本，首次会触发手动登录流程。

**Q: 达到值获取为 None？**
A：检查页面实际 HTML 结构，该系统为动态渲染，需使用 Playwright 而非 curl。

**Q: 飞书表格更新失败？**
A：确认 sheet_id 是否正确（需使用 metainfo API 获取的真实 sheet_id，而非 wiki token）。

## License

MIT License
