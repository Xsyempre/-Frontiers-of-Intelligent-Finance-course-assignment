# 子任务二：AI辅助分类财务舞弊 – 模型标注（Pilot 阶段）

你正在参与一项学术课程作业。我将上传以下文件：
1. `作业说明_update.html`（完整作业规范，含字段定义、判断标准、正反例）
2. `minimax_ver.xlsx`（包含第2组5%抽样样本，含原始 Activity 列和六个待填充的空字段列）

请你扮演一个**严谨的会计学AI助手**，严格按照作业说明中的规则，对 minimax_ver.xlsx 中的**每一行**进行标注，填充相应字段。

---

## 一、输出要求

- 更新该excel。文件名仍为"minimax_ver.xlsx"。
- 必须包含以下列：
  - Activity（原始违规描述，原样复制）
  - ann_related（0 或 1）
  - ann_year（JSON 数组字符串，如 "[2017,2018]"；若无则填 `null`）
  - ann_fin_flag（0 或 1 或 `null`）
  - ann_fin_info（JSON 数组字符串，如 '[{"year":2017,"elements":["收入","资产"]}]'；若无则填 `null`）

## 二、字段填写规则（严格执行）

请仔细阅读 `作业说明_update.html` 中以下章节：
- **任务二**：财务信息与会计要素（ann_fin_flag, ann_fin_info），注意前提条件

## 判断规则
- 若 `ann_related = 0`：直接输出 `{"ann_fin_flag": null, "ann_fin_info": null}`
- 若 `ann_related = 1`：
  - **ann_fin_flag**：
    - 1 = 影响财务信息（即财务报表中的科目金额错误）
    - 0 = 仅非财务信息（如公司治理披露、重大事项描述，金额无误）
  - 若 **ann_fin_flag = 1**，进一步识别受影响的**会计要素**（仅限：`资产`、`负债`、`所有者权益`、`收入`、`费用`、`利润`）。
  - **年份与要素必须一一对应**：不同年份分别列出，不能合并。

**关键约束**：
1. `ann_related=0` 时，`ann_year` 必须为 `null`，且后续 `ann_fin_flag`、`ann_fin_info`、`third_party_flag`、`third_party_list` 全部为 `null`。
2. `ann_related=1` 但 `ann_fin_flag=0` 时，`ann_fin_info` 为 `null`，且 `third_party_flag` 和 `third_party_list` 也为 `null`。
3. `ann_related=1` 且 `ann_fin_flag=1` 时，才能判断 `third_party_flag`，否则 `third_party_flag` 为 `null`。
4. 会计要素仅限六个：`资产`、`负债`、`所有者权益`、`收入`、`费用`、`利润`。禁止出现具体科目名（如“应收账款”）。
5. 年份与要素必须一一对应：若不同年份影响不同要素，请在 `ann_fin_info` 中按 `year` 分开记录。

## 三、示例

（以下示例仅用于说明格式，实际以作业说明为准）

**输入 Activity**：
> 公司在2017年年度报告中虚增营业收入3000万元，同时虚增应收账款5000万元。上述行为由供应商北京宏远商贸有限公司配合签订虚假合同实现。

**输出excel文件**：
| Activity | ann_related | ann_year | ann_fin_flag | ann_fin_info |
|----------|-------------|----------|--------------|--------------|
| 公司在2017年年度报告中虚增营业收入3000万元，同时虚增应收账款5000万元。上述行为由供应商北京宏远商贸有限公司配合签订虚假合同实现。 | 1 | [2017] | 1 | [{"year":2017,"elements":["收入","资产"]}] |

