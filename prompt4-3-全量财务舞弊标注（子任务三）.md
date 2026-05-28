# 子任务三：全量财务舞弊标注 – 第三方配合造假判断

你将继续担任会计学AI助手。我将上传一个 Excel 文件，该文件已经包含 `Activity`, `ann_related`, `ann_year`, `ann_fin_flag`, `ann_fin_info` 字段（由子任务一和子任务二生成）。请你严格按照《作业说明》中的规则，在现有文件基础上**增加两个新字段**：`third_party_flag` 和 `third_party_list`。

## 输入文件
`STK_labeled_夏思远_2023110537_G02.xlsx`  
路径：`D:\5th My Projects\FOR TEST\全量数据标注\STK_labeled_夏思远_2023110537_G02.xlsx`  
该文件已包含列：`Activity`, `ann_related`, `ann_year`, `ann_fin_flag`, `ann_fin_info`

## 输出要求
- **更新**同一个 Excel 文件，文件名和路径不变。
- 文件必须包含以下列（顺序固定）：
  1. `Activity`
  2. `ann_related`
  3. `ann_year`
  4. `ann_fin_flag`
  5. `ann_fin_info`
  6. `third_party_flag`（0 或 1 或 `null`）
  7. `third_party_list`（JSON 数组字符串，如 '[{"name":"X公司","type":"供应商","role":"虚假合同"}]'；否则 `null`）

- **null 统一用小写字符串 `null`**
- 行数不变，顺序不变

## 判断规则（严格按作业说明，仅任务三）

请仔细回顾 `作业说明_update.html` 中**任务三**的内容。

- **前提条件**：只有当 `ann_related = 1` 且 `ann_fin_flag = 1` 时，才判断 `third_party_flag` 和 `third_party_list`；否则两者均为 `null`。
- **第三方定义**：上市公司本身（及子公司）、实际控制人**除外**。包括客户、供应商、银行/金融机构、中介机构（券商、会计师、评估机构）、其他企业、自然人。
- **配合定义**：必须存在**主动协助造假行为**（虚假合同、资金回流、出借账户、代持股份、伪造流水等）。正常业务往来、被动配合调查、仅作为交易对手出现，都不算。
- 若存在：
  - `third_party_flag = 1`
  - `third_party_list` 为数组，每项包含：
    - `name`：尽量提取完整法定名称；若只有简称则原样提取，并在 `role` 中注明“文本中未出现全称”
    - `type`：从【客户、供应商、银行/金融机构、券商/保荐机构、会计师事务所、评估机构、自然人、其他企业】中选择
    - `role`：简要描述具体配合行为
- 若不存在：`third_party_flag = 0`，`third_party_list = null`

## 关键约束
- 不要修改已有的字段
- 配合行为需明确且有主动协助证据

## 工作流程
1. 读取指定路径的 Excel 文件
2. 根据每行的 `ann_related`、`ann_fin_flag` 和 `Activity` 判断 `third_party_flag` 和 `third_party_list`
3. 将这两个新列添加到现有数据中，覆盖保存原文件

## 注意
- 只增加两个字段
- 全量数据，请高效处理

现在请开始执行。