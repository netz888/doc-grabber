# 在线文档运行时导出工具（doc-grabber）

> **腾讯文档导出 / 飞书文档导出**工具 —— 把腾讯文档在线表格、飞书多维表格（Bitable）一键导出为本地 Excel（`.xlsx`）。基于 Python + Playwright 的浏览器运行时导出方案。
>
> 关键词：腾讯文档导出、飞书文档导出、腾讯文档下载、飞书多维表格导出、在线表格导出 Excel、Tencent Docs export、Feishu Bitable export。

这个项目包含**两个独立的命令行工具**，用于把在线文档表格导出为本地 Excel 文件（`.xlsx`）：

| 工具 | 文件 | 适用平台 | 是否需要登录 |
|------|------|----------|--------------|
| 腾讯文档导出器 | `extract_qq_sheet_runtime.py` | 腾讯文档在线表格 | 否（公开文档） |
| 飞书多维表格导出器 | `extract_feishu_base_runtime.py` | 飞书多维表格（Base/Bitable） | 是（手动登录） |

和很多“抓接口解析压缩数据”的方案不同，本项目采用的是**浏览器运行时导出**：用 Playwright 启动 Chromium 打开页面，直接读取页面前端运行时里的数据对象，再用 openpyxl 写成 Excel。这样导出的数据更接近你在云端真正看到的内容。

---

## ⚠️ 免责声明

- 本项目**仅供个人学习与技术研究使用**，用于了解浏览器自动化、前端运行时数据读取和 Excel 生成等技术。
- 使用者应当**遵守所在国家/地区的法律法规**，以及目标平台（腾讯文档、飞书等）的用户协议与服务条款。
- 请**仅对你拥有合法访问权限的文档**使用本工具，**不得用于抓取、复制或传播他人享有权利或涉及隐私、商业秘密的数据**。
- **因使用本项目所产生的任何后果（包括但不限于违反平台条款、侵犯他人权益、法律责任等），均由使用者自行承担，与本项目作者无关。**
- 如果你不同意以上条款，请不要使用本项目。

---

## 工具一：腾讯文档导出器

### 它是怎么工作的

- 用 Playwright 启动自带的 Chromium
- 打开腾讯文档页面
- 直接读取页面运行时里的 `window.SpreadsheetApp.workbook`
- 按浏览器真实显示层数据导出 Excel

支持多 sheet 工作簿，并尽量保留：超链接、合并单元格、行高、列宽、冻结窗格、基础字体/填充/边框/对齐。

### 适用范围

- 腾讯文档在线表格，URL 类似 `https://docs.qq.com/sheet/...`
- 文档可以在浏览器里正常打开（公开可访问、无需登录）
- 页面加载后能访问运行时对象 `window.SpreadsheetApp.workbook`

### 如何启动

```powershell
python extract_qq_sheet_runtime.py
```

程序是**纯命令行交互模式**，启动后会依次引导你：

1. **输入腾讯文档 URL**

   ```text
   请输入腾讯文档 URL:
   https://docs.qq.com/sheet/XXXXXXXXXXXXXXXX?tab=xxxxxx
   ```

   如果 URL 为空或不像腾讯文档 URL，程序会要求重新输入。

2. **选择要导出的 sheet**（多个可见 sheet 时）

   ```text
   检测到多个 sheet，请选择要导出的编号：
   0. 全部
   1. Sheet1
   2. Sheet2
   请输入编号（如 0 或 1,3,4）:
   ```

   - `0` 导出全部**可见** sheet（隐藏 sheet 不会自动进入范围）
   - `1,3,4` 导出指定编号
   - 只有 1 个 sheet 时直接导出

3. **确认输出目录**

   ```text
   默认输出目录名: exports_20260417_021312
   回车使用默认目录名，或输入自定义目录名:
   ```

   直接回车用默认名，或输入自定义目录名。目录已存在时会拒绝并要求重新输入，避免覆盖。

4. **自动导出**，完成后打印输出目录、Excel 路径、manifest 路径和导出的 sheet 列表。

---

## 工具二：飞书多维表格导出器

### 它是怎么工作的

飞书多维表格需要登录，所以程序会启动**可见浏览器（非无头）**，让你在弹出的窗口里手动完成飞书登录，检测到登录完成后再提取数据。

- 通过 Playwright 执行 JavaScript，读取飞书前端运行时对象：
  - `window.store.getState().bitable.Tables` → 表格列表
  - `window.store.getState().bitable.Fields[tableId]` → 字段定义
  - `window.bitableStore.base.getTableById(tableId).getRecordIds()` → 记录 ID 列表
  - `window.bitableStore.getCellValue(tableId, recId, fieldId)` → 单元格值
- **不依赖飞书开放平台 API，不需要创建飞书应用**
- 会把 option ID 解析为选项名称、把时间戳转为可读日期
- 多表格写成多 sheet 的 Excel 工作簿

### 适用范围

- 飞书多维表格，URL 形如 `https://xxx.feishu.cn/base/<token>?table=<tableId>&view=<viewId>`
- 文档需要你自己的飞书账号登录后才能访问

### 如何启动

```powershell
python extract_feishu_base_runtime.py
```

交互流程：

1. 输入飞书多维表格 URL
2. 程序启动可见浏览器并打开 URL（会跳转到登录页）
3. **在弹出的浏览器窗口里手动完成飞书登录**
4. 登录完成后程序自动读取表格列表
5. 多表格时按编号选择，单表格直接导出
6. 自动导出为 Excel，并写出 `tables_manifest.json`

---

## 环境要求

两个工具的环境要求相同。

### 1. Python

建议 Python 3.10 或更高版本：

```powershell
python --version
```

### 2. Python 依赖

```powershell
python -m pip install playwright openpyxl
python -m playwright install chromium
```

脚本默认使用 **Playwright 自带的 Chromium**，不依赖系统额外安装的 Google Chrome，程序会自动识别其安装路径，无需手动配置。

你可以这样检查 Python 是否能识别到 Playwright Chromium：

```powershell
@'
import extract_qq_sheet_runtime as m
print(m.detect_playwright_chromium_executable_path())
'@ | python -
```

能打印出一个 `.exe` 路径就说明浏览器已安装好。

### 3. 网络环境

- 能访问目标平台（`docs.qq.com` 或 `feishu.cn`）
- 当前工作目录有写权限

---

## 输出结果说明

每次导出目录里通常生成两个文件：

- **`document.xlsx`** — 主输出文件，包含选中的 sheet/表格、单元格值，以及尽量保留的结构信息（超链接、合并单元格、行高、列宽、冻结窗格、基础样式）。
- **`*_manifest.json`** — 导出元信息（表名、行列数、记录数等）。

> 注意：导出的 Excel 不保证和原文档 100% 一致，重点是**数据正确、基本结构正确**，某些高级表现可能存在差异。

---

## 运行测试

```powershell
python -m unittest -v
```

测试覆盖：sheet 选择解析、非法输入处理、默认输出目录名、已存在目录拒绝逻辑、隐藏 sheet 过滤、字段/记录解析、workbook 导出结构等。

---

## 项目文件说明

- `extract_qq_sheet_runtime.py` — 腾讯文档导出器（主程序）
- `extract_feishu_base_runtime.py` — 飞书多维表格导出器（主程序）
- `test_extract_qq_sheet_runtime.py` — 腾讯文档导出器测试
- `test_extract_feishu_base_runtime.py` — 飞书多维表格导出器测试

---

## 常见问题

**只装 Python 就够了吗？**
不够。还需要 `playwright`、`openpyxl`，并执行 `python -m playwright install chromium`。

**启动失败提示找不到浏览器？**
多半是没执行 `python -m playwright install chromium`。先用上面的检测命令确认能打印出 Chromium 路径。

**`0=全部` 为什么漏了某些 sheet？**
腾讯文档导出器里 `0` 只包含**可见** sheet，隐藏 sheet 不会默认导出（有意设计，避免隐藏页导致加载问题）。

**程序会覆盖已有目录吗？**
不会，目录已存在时会拒绝并要求重新输入。

**飞书导出为什么要弹出浏览器？**
因为飞书多维表格需要登录，必须由你本人在浏览器里手动登录，程序不会、也无法替你保存或绕过登录凭据。
