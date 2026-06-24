# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Monorepo 说明

本文件夹 `dir log/` 是 `manufacturing_process_schedule/` monorepo 中的一个独立子项目，记录了 DIR 工作簿的历史版本。**不要读取同级目录（`dir/`、`oe/` 等）的文件**，每个子目录是独立的项目。

## 项目概述

这是 **DIR (Drawing Inspection Report) 工作簿**的旧版 VBA 代码存档，以及配套的 PowerShell/AHK/JS 自动化脚本。

核心工作簿：`Measuring Tools.xlsm`。`.bas`/`.cls` 文件是从工作簿导出的 VBA 模块，用于版本控制。修改后需在 VBA IDE 中手动导入（File → Import File）。

## 工作表格式
A - Drawing Number
B - Revision
C - Job Number
D - File Name (生成的 DIR 文件的文件名)
E - OE Number
F - PO Number
G - Description
H - Customer
I - Start Time
J - Finish Time
L - Date

## 主要功能与工作流

### DIR 文件创建流程（核心业务流）

用户在工作表 **C 列**输入 Job Number → `mod_Sheet.bas` 捕获 `Worksheet_Change` → `FindDrawingNumbers` 查询 SQLite 数据库 → 弹出 `JobSelector` 表单 → 用户选择图纸编号 → `EventHandler_JobSelectorButton` 将数据写回工作表并构造 D 列公式。

之后，用户点击操作区按钮：
- **创建 DIR**：`mod_CheckSelectedCell.bas` 的 `CheckSelectedCell` — 读取当前行 D 列的文件名，将 `DIR Template.xlsm` 复制到网络路径并写入表头数据
- **打开 DIR**：`mod_OpenFile.bas` 的 `OpenFile` — 根据同样路径规则打开已有的 `.xlsm` 文件
- **打开 PDF**：`mod_OpenPDF.bas` 的 `OpenPDFFileInCompanyFolder` — 依次尝试 Bubble Drawing 子目录→基础路径（新版文件名）→基础路径（旧版文件名）

### 网络路径与客户路由

| Column H 值 | 目标路径 |
|-------------|---------|
| `Candu` | `\\rtdnas2\QCReports\FINAL REPORTS\CANDU  ENERGY\{col F}\` |
| `ATS` | `\\rtdnas2\QCReports\FINAL REPORTS\ATS  Energy\` |
| `Kinectrics` | `\\rtdnas2\QCReports\FINAL REPORTS\KINECTRICS INC\` |

仅 Candu 客户且 F 列不为空时，会在基础路径下创建以 F 列值命名的子目录。

## 工作表列含义（DIR 跟踪表）

| 列 | 内容 |
|----|------|
| A | 图纸编号（drawing number） |
| B | 版本号（revision） |
| C | Job Number（触发数据库查询） |
| D | 合成文件名公式：`=IF(AND(A="",B="",C=""),"",A&" Rev. "&B&" @"&C)` |
| E | OE Number |
| F | PO Number（清洁后，去除修订后缀） |
| G | 描述（来自 assemblies 表） |
| H | 客户名称（用于路由） |

## 数据库

- **路径**：`\\rtdnas2\OE\jobs.db`（SQLite）
- **驱动**：`mod_SQLite.bas` 封装 `SQLite3_StdCall.dll`（需与工作簿同目录）
- **表 `jobs`**：`job_number`, `part_number`, `oe_number`, `po_number`, `part_description`
- **表 `assemblies`**：`part_number`, `drawing_number`, `description`

查询逻辑：job_number → `jobs` 获取 part_number → `assemblies` 获取关联的所有 drawing_number。

## 模块说明

| 文件 | 职责 |
|------|------|
| `mod_PublicData.bas` | 全局状态：`lastEditedRow`（记录最后编辑的行号） |
| `mod_Sheet.bas` | Worksheet_Change：C 列变化时触发 `FindDrawingNumbers` |
| `mod_FindDrawingNumbers.bas` | 查询数据库，加载 JobSelector 表单 |
| `EventHandler_JobSelectorButton.bas` | JobSelector OK 按钮处理：写入 A/E/F/G 列，D 列写入公式；用正则提取标准 PO 格式（`RT\d{2}-\d{5}-PN-R\d{3}`） |
| `mod_CheckSelectedCell.bas` | 复制 DIR 模板到网络路径，写入表头信息（H7/H8/C8/C6/H6/C7/C11） |
| `mod_OpenFile.bas` | 打开已有 DIR `.xlsm` 文件 |
| `mod_OpenPDF.bas` | 查找并打开对应 PDF（多路径回退策略） |
| `mod_GenerateString.bas` | `GenerateDiscusString`：生成 Discus 用字符串；`GeneratePDFStringValue`：生成 PDF 文件名（格式：`{A} Rev{B} {G Title Case}-ballooned`） |
| `mod_SQLite.bas` | SQLite3 VBA 封装 |
| `Measuring Tool.bas` | `mod_CheckSelectedCell.bas` 的早期版本（使用本地路径），仅供参考 |

## PowerShell 脚本

| 脚本 | 用途 |
|------|------|
| `Excel to PDF.ps1` | 将当前目录所有 Excel 文件转换为 PDF |
| `Excel to PDF Combined.ps1` | 转换并合并为单个 PDF |
| `Combine Excel to PDF by DIR.ps1` | 扫描 `DIR 1`/`DIR 2`... 子目录，每个目录内 Excel→PDF 后合并；**需要 Adobe Acrobat Pro** |
| `Merge Measurement Results.ps1` | 合并 Faro CMM 导出的 CSV；按文件名中的 `(1)`/`(2)` 或 `-1`/`-2` 排序；支持单区域/多区域（`<No. # area>` 标头）两种格式；数值格式化为 4 位小数 |
| `Merge Faro.ps1` | 更简单的 Faro 合并：以 `-1.csv` 为基准，将 Part 列复制到所有文件，按 Part 列排序后合并 |

运行 PowerShell 脚本：
```powershell
# 在目标目录下执行
.\Excel to PDF.ps1
.\Merge Measurement Results.ps1
powershell -File "Combine Excel to PDF by DIR.ps1" -ParentFolder "C:\path\to\folder"
```

## 其他文件

- `align.ahk` / `NumPadPlusToCtrlShiftB.ahk`：AutoHotkey 快捷键脚本
- `pdfMerge.js`：Node.js PDF 合并脚本（备用方案）
- `GenerateRandomNumber*.bas` / `GenRandomNumber_Click*.bas`：随机数生成（含 backup 备份版本）
- `Form_BatchInput/`：旧版批量输入表单代码目录
