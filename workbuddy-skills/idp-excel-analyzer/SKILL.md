---
name: idp-excel-analyzer
description: 用于对个人发展计划表（IDP）Excel 文件进行四维度分析：基本情况、IDP 目标、行动计划、目标周期，并生成 Excel 分析表与 Markdown 报告。当用户要求分析 IDP / 个人发展计划 / 员工发展计划 / 人才发展计划等 Excel 文件时触发。支持中文和英文输出。
agent_created: true
---

# IDP Excel 分析器

## 概述

读取个人发展计划表（IDP）Excel 文件，自动完成数据清洗、四维度分析与可视化，输出 Excel 分析表和 Markdown 分析报告。适用于 HR、零售、门店管理等领域的人才发展计划分析。支持中文（`analyze_idp.py`）和英文（`analyze_idp_en.py`）两种输出语言。

## 何时使用

当用户请求满足以下任一条件时触发本技能：
- 分析个人发展计划表 / IDP / 员工发展计划 / 人才发展计划
- 对 Excel 格式的 IDP 数据按维度分析
- 需要生成 IDP 统计报告、仪表盘或汇总表
- 用户要求用英文输出分析结果时，使用 `analyze_idp_en.py`

## 分析维度

1. **基本情况分析**：员工总数、职位分布、区域分布、店铺分布、经理下属分布、跨城市/跨部门调动意愿、填写日期分布。
2. **IDP 目标分析**：目标填写完整度、目标类型分布、技能属性（软性/硬性）分布、主题分类、目标描述关键词提取。
3. **行动计划分析**：70-20-10 学习模型填写率、常见行动方式统计、资源需求分析、员工 IDP 完整度评分。
4. **目标周期分析**：总目标达成时间、中期/终期回顾时间分布、周期长度统计、时间逻辑异常检查。

## 工作流程

1. **读取数据**：使用 pandas 读取 Excel 第一个工作表，跳过提示行，以第 3 行为表头。
2. **数据清洗**：
   - 去除文本字段首尾空格
   - 统一职位名称大小写/中英文变体（如 SSA/ssa/主管/Supervisor/supervisor 等归一化）
   - 将中文日期格式（如“2026年12月31日”）解析为 datetime
   - 标记缺失的目标类型与行动计划
3. **分析计算**：
   - 对目标类型提取技能属性（软性/硬性/领导力）与主题
   - 对 70/20/10 行动计划字段进行覆盖度统计
   - 对行动方式字段提取关键词并计数
   - 使用 jieba 对目标描述与资源字段进行中文关键词提取
   - 计算每位员工的完整度得分（0-12 分）
   - 计算时间周期与异常标记
4. **生成图表**：使用 matplotlib 生成职位分布、目标主题、技能属性、行动计划、完整度、时间线等图表。
5. **输出交付物**：
   - `个人发展计划表分析_结果.xlsx`：包含原始数据、基本情况汇总、目标分析、行动计划分析、周期分析、员工明细评分等工作表。
   - `个人发展计划表分析_报告.md`：结构化分析报告，含执行摘要、四维度分析、洞察与建议。
   - `charts/` 目录：分析图表 PNG 文件。

## 执行脚本

复杂分析可通过脚本自动执行：

- `scripts/analyze_idp.py`：中文输出，输入 Excel 路径，输出 Excel 分析表、Markdown 报告和图表。
- `scripts/analyze_idp_en.py`：英文输出，功能与中文版相同，适用于需要英文报告的场景。支持额外的"区域"（Region）列。

运行方式：

```bash
# 中文输出
python scripts/analyze_idp.py <input_excel_path> [output_dir]

# 英文输出
python scripts/analyze_idp_en.py <input_excel_path> [output_dir]
```

若未提供 `output_dir`，默认保存到输入文件所在目录的 `Final` 子目录。

英文版输出文件名：
- `IDP_Analysis_Results.xlsx`
- `IDP_Analysis_Report.md`
- `charts/` 目录：7 张分析图表（含区域分布图）

## 依赖

- Python 3.10+
- pandas
- openpyxl
- matplotlib
- jieba
- tabulate

## 数据质量处理规则

- 目标类型为空：标记为“未填写”。
- 行动计划为空：标记为“未填写”。
- 职位名称归一化：SSA/ssa → SSA；SA/sa/SALES/Senior sales → SA；Supervisor/supervisor/SPV/主管 → Supervisor；Senior Admin/Senior admin → Senior Admin。
- 日期异常：总目标达成时间早于填写日期、中期晚于终期、终期晚于总目标的情况单独标记。

## 注意事项

- 只读取 Excel 文件，不修改原始文件。
- 第一个工作表应为 IDP 主数据表，列名需包含"工号"、"员工姓名"、"经理姓名"、"目前职位"、"店铺"、"日期"、"发展目标1/2"相关字段及时间字段。
- 英文版脚本（`analyze_idp_en.py`）额外支持"区域"（Region）列，位于"目前职位"和"店铺"之间。
- 若目标类型格式与预期不同，可能需要调整 `normalize_position` 和 `classify_goal_theme` 函数。
- 英文版将中文调动意愿（是/否）自动翻译为 Yes/No。
