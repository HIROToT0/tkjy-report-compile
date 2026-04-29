# 太科检测报告全流程自动化

> TKMgt 任务数据抓取 → 飞书表格 → 报告编制系统自动化（判断状态、编辑样品达到值、生成上传报告、回填状态）

[![Python 3](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-自动化-orange.svg)](https://playwright.dev/)

## 功能特性

- ✅ **TKMgt 数据抓取**：从 TKMgt 检测系统批量提取任务数据，写入飞书表格
- ✅ **自动判断报告状态**：跳过非待编制报告，减少无效操作
- ✅ **批量编辑样品达到值**：支持多组样品并行处理
- ✅ **异常样品自动识别**：达到值 < 100 或显示"无效"自动标记
- ✅ **报告自动生成与上传**：保存 → 生成报告 → 上传报告一键完成
- ✅ **状态回填**：处理结果实时写入飞书表格（报告提交状态 / 异常情况说明）

## 完整流程

```
TKMgt检测系统
    │  抓取任务数据（reportCode/productName/entrustDate/requestTestDate...）
    ▼
飞书表格（任务清单）
    │  读取报告编号
    ▼
报告编制系统（comp1.taikotest.com）
    │  判断状态 → 编辑样品达到值 → 异常判定
    │  生成报告 → 上传报告
    ▼
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

```python
# TKMgt 检测系统
TKMGT_PHONE = "***"          # 填入实际手机号
TKMGT_COOKIE_FILE = "/tmp/tkmgt_full_cookies.txt"

# 报告编制系统
REPORT_USERNAME = "***"      # 填入实际账号
REPORT_PASSWORD = "***"      # 填入实际密码
REPORT_COOKIE_FILE = "~/.hermes/report_cookies_compile.txt"
REPORT_ID = "291"            # 混凝土试块抗压模板ID

# 飞书应用
FEISHU_APP_ID = "***"        # 填入实际 app_id
FEISHU_APP_SECRET = "***"    # 填入实际 app_secret
FEISHU_SPREADSHEET_TOKEN = "***"   # 填入实际 spreadsheet token
FEISHU_SHEET_ID = "***"      # 填入实际 sheet_id
```

### 运行

```bash
# 完整流程（TKMgt抓取 → 飞书 → 报告编制）
python tkjy_full_automation.py
```

## 核心判定规则

| 条件 | 操作 |
|------|------|
| 达到值 ≥ 100 | 点击"确定"，正常提交 |
| 达到值 < 100 | 点击"取消"，标记"异常" |
| 显示"无效" | 点击"取消"，标记"异常" |
| 多组样品任一异常 | 整体标记"异常"，跳过剩余组 |

## 关联系统

| 系统 | 地址 |
|------|------|
| TKMgt 检测系统 | http://test.tkjy.com/TKMgt/ |
| 报告编制系统 | http://comp1.taiketest.com/Report/ |
| 飞书表格（任务清单） | https://ccnlg9zq6b6x.feishu.cn/wiki/WTMWwhfcxiKjvkkLKhFc3haAnOf |

## 常见问题

**Q: TKMgt 接口返回空数据？**
A: Cookie 过期，需重新登录 TKMgt。

**Q: 页面停留在登录页，无法自动登录？**
A: 删除 `~/.hermes/report_cookies_compile.txt`，重新运行脚本，首次会触发手动登录。

**Q: 达到值获取为 None？**
A：该系统为动态渲染，需使用 Playwright 而非 curl。

**Q: 飞书表格更新失败？**
A：确认 sheet_id 正确（需使用 metainfo API 获取的真实 sheet_id，而非 wiki token）。

## License

MIT License
