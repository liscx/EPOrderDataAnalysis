---
name: order-data-analysis
description: 对商品交易订单数据进行自动化分析，包含数据清洗、供应商/采购商维度分析、专区分析及核心指标提取。支持通过命令行参数指定输入输出路径。
triggers:
  - 分析订单 Excel 数据
  - 执行订单数据清洗和汇总
  - 统计采购人与供应商排行
  - 提取核心话术指
  - 订单数据
requires:
  bins: [python]
---

# 订单数据分析集成技能 (Order Data Analysis)

此技能用于驱动 `EPOrderDataAnalysis` 系统执行全自动化的订单数据分析工作流。

## 场景说明
当你需要对包含“采购企业”、“供应商”、“订单金额”等字段的原始交易 Excel 文件进行深度分析时使用。该技能会自动从你的**下载文件夹**中寻找最新的“阳光优采交易订单”文件，并将分析结果产出到**桌面**。

## 自动化执行流程 (推荐)

运行以下 PowerShell 命令，它将自动定位最新下载的订单文件并开始分析：

```powershell
# 1. 自动寻找下载文件夹中最新的“阳光优采交易订单”文件
$inputFile = Get-ChildItem -Path "$HOME\Downloads\阳光优采交易订单*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName

# 2. 检查文件是否存在
if (-not $inputFile) { Write-Error "未在下载文件夹中找到匹配的订单文件。"; return }

# 3. 设置输出路径为桌面
$outputFile = "$HOME\Desktop\分析结果_$((Get-Date).ToString('yyyyMMdd_HHmm')).xlsx"

# 4. 执行分析
python main.py --input "$inputFile" --output "$outputFile"

# 5. 完成提示
Write-Host "分析完成！结果已保存至桌面: $outputFile"
```

## 指令规范

### 1. 全量自动化分析
指令：`分析最新的订单数据并发送结果`
技能逻辑：
1.  **动态路径获取**：自动执行 PowerShell 脚本（见上方）获取最新输入文件，并确定桌面输出路径。
2.  **时间区间解析**：
    *   如果用户提到特定月份（如“3月订单”）：
        *   若该月为**过去月份**，区间为该月 1 日至该月最后一天。
        *   若该月为**当前月份**（如 4 月 14 日询问 4 月数据），起始日期为该月 1 日，**结束日期为今天**（即 2026-04-14）。
    *   将计算出的时间区间通过 `--start` 和 `--end` 参数传递给 `main.py`。
3.  **脚本执行**：运行 `python main.py --input "$inputFile" --output "$outputFile" --start "<计算的开始日期>" --end "<计算的结束日期>"`。
4.  **结果交付**：分析完成后，必须将生成的桌面文件上传并发送给用户。

### 2. 指定参数分析
```powershell
python main.py --input "<路径>" --output "<路径>" --start "YYYY-MM-DD" --end "YYYY-MM-DD"
```

## 输入参数
- `input`: (自动/可选) 原始数据 Excel 的绝对路径。
- `output`: (自动/可选) 分析结果保存的绝对路径。
- `start`: (可选) 分析起始日期，格式为 `YYYY-MM-DD`。
- `end`: (可选) 分析结束日期，格式为 `YYYY-MM-DD`。

## 依赖环境
确保 Python 运行环境中已安装：`pandas`, `openpyxl`, `pyyaml`。
