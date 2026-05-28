# 子任务二：全量财务舞弊标注 – 财务信息与会计要素判断

你将继续担任会计学AI助手。我将上传一个 Excel 文件，该文件已经包含 `Activity`、`ann_related`、`ann_year` 字段（由子任务一生成）。请你严格按照《作业说明》中的规则，在现有文件基础上**增加两个新字段**：`ann_fin_flag` 和 `ann_fin_info`。

## 输入文件
`STK_labeled_夏思远_2023110537_G02.xlsx`  
路径：`D:\5th My Projects\FOR TEST\全量数据标注\STK_labeled_夏思远_2023110537_G02.xlsx`  
该文件已包含列：`Activity`, `ann_related`, `ann_year`

## 输出要求
- **更新**同一个 Excel 文件，文件名和路径不变。
- 文件必须包含以下列（顺序固定）：
  1. `Activity`
  2. `ann_related`
  3. `ann_year`
  4. `ann_fin_flag`（0 或 1 或 `null`）
  5. `ann_fin_info`（JSON 数组字符串，如 '[{"year":2017,"elements":["收入","资产"]}]'；否则 `null`）

- **null 统一用小写字符串 `null`**
- 行数不变，顺序不变

## 判断规则（严格按作业说明，仅任务二）

请仔细回顾 `作业说明_update.html` 中**任务二**的内容。

- 若 `ann_related = 0`：`ann_fin_flag = null`，`ann_fin_info = null`
- 若 `ann_related = 1`：
  - `ann_fin_flag`：
    - 1 = 影响财务信息（即财务报表中的科目金额错误）
    - 0 = 仅非财务信息（如公司治理披露、重大事项描述，金额无误）
  - 若 `ann_fin_flag = 1`，进一步识别受影响的**会计要素**（仅限：`资产`、`负债`、`所有者权益`、`收入`、`费用`、`利润`）
  - **年份与要素必须一一对应**：不同年份分别列出，不能合并。例如：`[{"year":2016,"elements":["收入","资产"]},{"year":2017,"elements":["收入","资产","费用","利润"]}]`

## 关键约束
- `ann_related = 1` 但 `ann_fin_flag = 0` 时，`ann_fin_info` 为 `null`
- 禁止使用具体科目名（如“应收账款”）

## 工作流程
1. 读取指定路径的 Excel 文件
2. 根据每行的 `ann_related` 和 `Activity`，计算 `ann_fin_flag` 和 `ann_fin_info`
3. 将这两个新列添加到现有数据中，覆盖保存原文件

## 注意
- 只增加两个字段，不要修改已有的 `Activity`、`ann_related`、`ann_year` 列
- 全量数据，请高效处理

现在请开始执行。