---
name: tkjy-report-compile
category: productivity
description: 操作太科检测报告编制系统（comp1.taiketest.com）完成混凝土试块抗压报告的批量编制——从飞书表格读取报告编号、判断状态、编辑样品达到值、生成上传报告、回填飞书表格状态。
---

# 太科检测报告编制系统自动化

## 简介

操作太科检测报告编制系统（comp1.taiketest.com）完成混凝土试块抗压报告的批量编制。核心流程：从飞书表格读取报告编号 → 判断报告状态 → 编辑样品达到值 → 生成并上传报告 → 回填飞书表格。

---

## 关联系统

| 系统 | 地址 | 用途 |
|------|------|------|
| 太科检测报告系统 | http://comp1.taiketest.com/Report/ | 报告编制 |
| 飞书表格（任务清单） | https://ccnlg9zq6b6x.feishu.cn/wiki/WTMWwhfcxiKjvkkLKhFc3haAnOf | 报告编号清单 |
| 飞书 API | open.feishu.cn | 读写表格状态 |

---

## 环境信息

- **账号**: `***`（填入实际账号）
- **密码**: `***`（填入实际密码）
- **Cookie 文件**: `~/.hermes/report_cookies_compile.txt`（首次需手动保存，后续复用）
- **飞书应用 CLI**: app_id=`***`, app_secret=`***`
- **报告系统 reportId**: `291`（混凝土试块抗压专用）

---

## 完整操作流程

### 第 0 步：环境准备

```bash
# 安装 Playwright（如果尚未安装）
pip install playwright
playwright install chromium

# 确保 cookie 文件存在
touch ~/.hermes/report_cookies_compile.txt
```

### 第 1 步：读取飞书表格，获取报告编号清单

```python
import subprocess, json

# 获取飞书 token
r1 = subprocess.run([
    'curl', '-sX', 'POST',
    'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
    '-H', 'Content-Type: application/json',
    '-d', json.dumps({
        "app_id": "***",      # 填入实际 app_id
        "app_secret": "***"   # 填入实际 app_secret
    })
], capture_output=True, text=True, timeout=15)
token = json.loads(r1.stdout)['tenant_access_token']

# 读取表格 A-F 列（前10行）
r = subprocess.run([
    'curl', '-s',
    'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/***/values/***!A1:F10',
    '-H', f'Authorization: Bearer {token}'
], capture_output=True, text=True, timeout=15)

data = json.loads(r.stdout)
rows = data['data']['valueRange']['values']
# rows[0] 是表头，rows[1:] 是数据
# 第0列=报告编号，第1列=检测对象，第4列=报告提交状态，第5列=异常情况说明
```

**避坑：**
- 飞书表格 range 要精确到具体 sheet_id（如 `ac04d9!A1:F10`），不能只写范围名或 spreadsheet token
- 首次读取前先确认 sheet_id 是否正确，错误则返回 `INVALID_PARAMETER`
- 多余额定要 `.get('data', {}).get('valueRange', {}).get('values', [])` 防御性读取
- 飞书 token 有 2 小时有效期，过期需重新获取

### 第 2 步：登录报告系统（强烈建议用 Playwright，不推荐 curl）

原因：
- 该站登录后设置多个 Cookie（latoToken + JSESSIONID），且有 HTTPOnly flag
- curl 无法正确接收和回传所有 Cookie，导致后续请求被重定向到登录页
- Playwright 能自动处理所有 Cookie 和会话

```python
from playwright.sync_api import sync_playwright

def login_to_report_system():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        
        # 访问登录页（末尾必须有斜杠）
        page.goto("http://comp1.taiketest.com/Report/")
        page.wait_for_load_state("networkidle")
        
        # 检查是否已登录
        if "/Report/" in page.url and "login" not in page.url.lower():
            print("已登录")
        else:
            page.fill('input[name="username"]', 'hewei')
            page.fill('input[name="password"]', '***')  # 填入实际密码
            page.click('button[type="submit"]')
            page.wait_for_load_state("networkidle")
        
        # 保存 Cookie 供后续复用
        context.storage_state(path="~/.hermes/report_cookies_compile.txt")
        print("登录成功，Cookie 已保存")
        return context, page
```

**避坑：**
- 登录 URL 是 `http://comp1.taiketest.com/Report/`，末尾有斜杠，遗漏会 302 重定向循环
- 登录成功后 URL 格式：`http://comp1.taiketest.com/Report/ptcompile/getCompile.do?reportId=291`
- 账号 hewei 在系统中有效，其他账号可能无 reportId=291 的权限
- Cookie 文件存储路径：`~/.hermes/report_cookies_compile.txt`（需绝对路径）

### 第 3 步：判断报告状态

进入报告页面后，获取"报告字段"内容：

```
URL格式：http://comp1.taiketest.com/Report/ptcompile/getCompile.do?reportId=291&reportCode={报告编号}
示例：http://comp1.taiketest.com/Report/ptcompile/getCompile.do?reportId=291&reportCode=THJHY20260000004268
```

```python
def check_report_status(page, report_code):
    url = f"http://comp1.taiketest.com/Report/ptcompile/getCompile.do?reportId=291&reportCode={report_code}"
    page.goto(url, timeout=30000)
    page.wait_for_load_state("networkidle")
    
    content = page.content()
    
    # 判断状态
    if "报告待编制" in content:
        return "待编制"
    elif "检验完成" in content:
        return "检验完成"
    elif "报告已上传" in content:
        return "已上传"
    else:
        # 兜底提取状态字段值
        import re
        match = re.search(r'报告字段.*?([^\s]{4,10})', content)
        return match.group(1) if match else "未知"
```

**避坑：**
- 页面是动态渲染（JavaScript），curl 拿不到实际内容，必须用 Playwright
- 如果页面长时间卡住，增加 `timeout=60000` 并减少等待
- 首次访问某报告编号时，如果系统无该编号数据，页面会显示空或跳回列表页
- 状态为"检验完成"时仍需继续操作（编辑样品 → 生成报告）

### 第 4 步：处理样品数据（核心逻辑）

#### 4.1 点击编辑按钮

页面下方显示样品编号列表（如 THJHY20260000004267-1），每个样品右侧有**编辑**按钮：

```python
def click_edit_for_sample(page, sample_index=0):
    """点击第 N 个样品的编辑按钮（0-based）"""
    edit_buttons = page.locator('button:has-text("编辑"), a:has-text("编辑")')
    count = edit_buttons.count()
    print(f"发现 {count} 个编辑按钮")
    
    if sample_index < count:
        edit_buttons[sample_index].click()
        page.wait_for_load_state("networkidle")
    else:
        raise Exception(f"没有第 {sample_index} 个样品")
```

**避坑：**
- 编辑按钮在动态渲染的表格中，需等待页面完全加载后再查找
- 弹窗出现后需等待 1-2 秒让弹窗完全渲染再操作
- 如果页面有多个编辑入口，确认点击的是样品行对应的按钮

#### 4.2 从弹窗获取"达到值"

```python
def get_reach_value_from_dialog(page):
    """从弹出框中提取达到值"""
    import re
    content = page.content()
    
    # 方法1：正则匹配"达到值"后面的数字
    match = re.search(r'达到值.*?(\d+\.?\d*)', content, re.DOTALL)
    if match:
        return float(match.group(1))
    
    # 方法2：直接查找数字（混凝土抗压通常50~200 MPa）
    numbers = re.findall(r'\d+\.?\d*', content)
    for num in numbers:
        val = float(num)
        if 50 <= val <= 200:
            return val
    
    return None
```

**避坑：**
- 弹窗中"达到值"字段可能包含单位（如"MPa"），读取时需去除非数字字符
- 如果正则匹配不到，尝试 `page.inner_text()` 逐行扫描
- 抗压强度达到值正常范围约 50~200 MPa，明显偏低（<100）需标记异常

#### 4.3 判定规则

```
如果 达到值 < 100  或者  显示"无效"：
    → 点击"取消"按钮
    → 飞书表格报告提交状态填"异常"，异常情况说明填"结果异常"
    → 跳到下一个报告编号

如果 达到值 ≥ 100：
    → 点击"确定"按钮提交
    → 继续处理下一组样品（如果有多组）
```

```python
def process_reach_value(page, reach_value):
    if reach_value is None:
        action = "取消"
    elif reach_value < 100:
        action = "取消"
    else:
        action = "确定"
    
    if action == "取消":
        page.click('button:has-text("取消")')
        return True  # 异常
    else:
        page.click('button:has-text("确定")')
        return False  # 正常
```

**避坑：**
- 样品有多组时（THJHY20260000004267-1、-2），**每一组都要处理**
- 多组样品中只要**任何一组**出现异常，整个报告编号就标记异常，跳过剩余样品
- 点击"确定"后弹窗关闭，需等待新页面渲染完成再进行下一步

### 第 5 步：生成报告并上传

#### 5.1 保存报告

```python
save_button = page.locator('button:has-text("保存"), a:has-text("保存")')
save_button.click()
page.wait_for_load_state("networkidle")
print("已点击保存")
```

#### 5.2 等待"生成报告"按钮就绪

```python
def wait_for_generate_button(page, timeout=60):
    """等待'生成报告'按钮可点击"""
    import time
    start = time.time()
    while time.time() - start < timeout:
        content = page.content()
        if "生成报告" in content:
            btn = page.locator('button:has-text("生成报告"), a:has-text("生成报告")')
            if btn.count() > 0 and btn.first.is_enabled():
                return True
        time.sleep(1)
    raise TimeoutError("生成报告按钮超时")
```

**避坑：**
- "生成报告"按钮在保存后需要一段时间才就绪（服务端处理），轮询等待时 sleep 不能太短（<0.5秒可能被限流）
- 超时时间至少设 60 秒

#### 5.3 点击生成报告 → 上传报告

```python
gen_btn = page.locator('button:has-text("生成报告"), a:has-text("生成报告")').first
gen_btn.click()
page.wait_for_load_state("networkidle")
print("已点击生成报告")

upload_btn = page.locator('button:has-text("上传报告"), a:has-text("上传报告")').first
upload_btn.click()
page.wait_for_load_state("networkidle")
print("已点击上传报告")
```

**避坑：**
- 上传报告后页面可能有确认提示，等待网络空闲后再进行下一个编号
- 如果上传失败（网络波动），整个流程需要从该报告编号重新开始

### 第 6 步：回填飞书表格

```python
def update_feishu_cell(token, row_index, col, value):
    """更新飞书表格单个单元格"""
    range_str = f"ac04d9!{col}{row_index}"
    payload = {
        "valueRange": {
            "range": range_str,
            "values": [[value]]
        }
    }
    r = subprocess.run([
        'curl', '-sX', 'PUT',
        'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/***/values',
        '-H', f'Authorization: Bearer {token}',
        '-H', 'Content-Type: application/json',
        '-d', json.dumps(payload)
    ], capture_output=True, text=True, timeout=15)
    result = json.loads(r.stdout)
    return result.get('code') == 0

# 使用示例
# E列=报告提交状态，F列=异常情况说明
# 第2行对应 rows[1]（第1行是表头）
update_feishu_cell(token, 7, 'E', "已提交")     # 正常提交
update_feishu_cell(token, 7, 'E', "异常")      # 异常情况
update_feishu_cell(token, 7, 'F', "结果异常")  # 异常说明
```

**避坑：**
- 飞书表格行号从 1 开始，`ac04d9!E2` = 第2行
- 更新频率不要太高（>5次/秒），建议每条间隔 0.5 秒
- 异常情况说明只在异常时填写，正常提交时不填

---

## 整体循环逻辑

```
for 每行报告编号 in 飞书表格:
    1. 判断是否已处理（报告提交状态 ≠ 空）
       → 已处理则跳过
    2. 进入报告页面 → 判断状态
       → 非"报告待编制" → 回填状态 → 下一个
    3. 依次点击每组样品的编辑按钮
       → 获取达到值
       → <100 或无效 → 取消 → 填"异常" → 跳下一个编号
       → ≥100 → 确定
    4. 所有样品处理完 → 保存 → 等生成报告 → 点生成 → 点上传
    5. 回填飞书"已提交"
```

---

## 常见错误处理

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| 页面停留在登录页 | Cookie 无效或过期 | 删除 cookie 文件，重新用 Playwright 登录 |
| `INVALID_PARAMETER` | sheet_id 或 range 错误 | 确认 sheet_id 不是 wiki_token，需用 metainfo API 获取真实 sheet_id |
| 编辑按钮点不到 | 页面未完全加载 | 增加 `wait_for_load_state("networkidle")` 或 `wait_for_timeout(2000)` |
| 达到值获取为 None | 正则匹配失败 | 检查页面实际 HTML 结构，尝试 `page.inner_text()` 逐行扫描 |
| 保存后按钮不出现 | 服务端处理中 | 增加等待时间，或刷新页面重试 |
| curl 返回空白 | Windows 路径含中文 | 用英文临时路径，避免 `桌面` 等中文目录 |
| 多个编辑按钮分不清 | 页面有多个编辑入口 | 通过按钮的 `nth()` 和附近文本（样品编号）定位唯一按钮 |

---

## 性能优化

- **Cookie 复用**：每处理完一个编号不要关闭浏览器 context，复用登录态直到所有编号完成
- **跳过已处理**：表格读取时过滤掉已提交/异常的编号，减少不必要的页面访问
- **并发处理**：多个报告编号之间完全独立，可以 `delegate_task` 并行处理（最多 3 个并发）
- **批量回填**：所有编号处理完后统一回填飞书（需记录每个编号的处理结果）

---

## 文件路径速查

| 文件 | 路径 |
|------|------|
| 报告系统 Cookie | `~/.hermes/report_cookies_compile.txt` |
| 报告系统登录页 | `http://comp1.taiketest.com/Report/` |
| 报告系统 reportId | `291`（混凝土试块抗压） |
| 飞书表格 spreadsheetToken | `***`（填入实际值） |
| 飞书表格 sheetId | `***`（填入实际值） |
| 飞书表格列映射 | A=报告编号, B=检测对象, C=检测日期, D=预计完成日期, E=报告提交状态, F=异常情况说明 |
