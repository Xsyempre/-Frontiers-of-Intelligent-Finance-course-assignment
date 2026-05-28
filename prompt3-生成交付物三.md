# 任务：生成交付物三 —— 人工标注与模型对比表（validation.xlsx）

我正在完成《智能财务前沿》课程作业。目前已完成 Pilot 阶段：
- 5% 抽样样本的人工基准标注
- 三个模型（DeepSeek、MiniMax、Mimo）分别对同一批样本进行了标注

现在需要你根据这些文件，生成一份 **validation.xlsx**，严格按照作业说明_update.html 中“交付物三”的要求。

## 一、输入文件（我将全部上传）

1. **作业说明_update.html**（规范依据）
2. **benchmark_G02.xlsx**（人工基准标注结果，含 Activity + 六个字段）
3. **deepseek_ver.xlsx**（DeepSeek 模型的标注结果）
4. **minimax_ver.xlsx**（MiniMax 模型的标注结果）
5. **mimo_ver.xlsx**（Mimo 模型的标注结果）

> 注：以上文件均基于同一份抽样样本，行数相同、行序相同（均以原始 Activity 为标识，可对齐）。

## 二、输出文件要求

- 文件名：`validation_夏思远_2023110537_G02.xlsx`
- 格式：`.xlsx`
- 内容必须包含以下信息（每行一条抽样记录）：

• 原始 Activity 文本（完整保留）
• 人工标注结果：ann_related、ann_fin_flag、third_party_flag 三个字段的人工判断值
• 模型 A 的预测结果（三个字段）
• 模型 B 的预测结果（三个字段）
• 模型 C 的预测结果（若使用三个模型）（三个字段）
• 每个模型在每个字段上的对比结果（TP / FP / FN / TN）
• 末行汇总各模型的 F1 Score（三个字段分别计算）

**注意**：对于 `ann_fin_flag` 和 `third_party_flag`，空值（前提不满足时）应统一表示为字符串 `"null"` （作业说明要求填 null，建议统一用 `null` 字符串）。

### 每个模型在每个字段上的对比结果（TP/FP/FN/TN）

你需要在表格**右侧**增加以下列，**逐条记录**每个模型在三个字段上的分类结果：

- 对于 `ann_related`：
  - `DeepSeek_ann_related_TP`（1/0）  
  - `DeepSeek_ann_related_FP`  
  - `DeepSeek_ann_related_FN`  
  - `DeepSeek_ann_related_TN`  
  （其他模型同理）

- 对于 `ann_fin_flag`（只在 `human_ann_related=1` 的子样本中计算，其他行此项可为空或标记为 `N/A`）：
  - `DeepSeek_ann_fin_flag_TP` … 等

- 对于 `third_party_flag`（只在 `human_ann_related=1` 且 `human_ann_fin_flag=1` 的子样本中计算，其他行标记为 `N/A`）

> **重要**：TP/FP/FN/TN 的计算规则：
> - TP：人工=1，模型=1
> - FP：人工=0，模型=1
> - FN：人工=1，模型=0
> - TN：人工=0，模型=0
> - 对于 `null` 值的处理：在有效子集内，人工为1或0（不会为null），所以不需要处理 null 比较。

### 末行汇总：各模型的 F1 Score

在表格的最后一行（所有数据行之后），汇总每个模型在三个字段上的 **F1 Score**，格式如：

| 模型 | ann_related F1 | ann_fin_flag F1 | third_party_flag F1 | 平均 F1 |
|------|----------------|------------------|----------------------|----------|
| DeepSeek | 0.80 | 0.75 | 0.60 | 0.72 |
| MiniMax | 0.90 | 0.85 | 0.80 | 0.85 |
| Mimo | 0.70 | 0.65 | 0.70 | 0.68 |

- F1 计算公式：`2 * (Precision * Recall) / (Precision + Recall)`，其中 Precision = TP/(TP+FP)，Recall = TP/(TP+FN)
- 对于某个字段如果有效样本数为0（例如没有 `third_party_flag` 正例），则 F1 填 `N/A`
- 平均 F1 只对有效数值求平均

## 三、工作步骤

1. 读取 `benchmark_G02.xlsx`，获取每条记录的 `Activity`、`ann_related`、`ann_fin_flag`、`third_party_flag`（作为人工基准）。
2. 依次读取三个模型的 Excel 文件，根据行序或 Activity 文本匹配，提取每个模型对同一条记录的三个字段预测值。
3. 按上述规则，逐行计算每个模型在每个字段上的 TP/FP/FN/TN（注意有效子集）。
4. 汇总每个模型的 F1 Score。
5. 生成一个 pandas DataFrame，按要求的列顺序输出到 `validation_姓名_学号_G02.xlsx`。
6. 对于无效子集（例如 `ann_fin_flag` 比较时，人工 `ann_related=0` 的行），TP/FP/FN/TN 列留空或填 `N/A`，F1 计算时忽略这些行。

## 四、补充要求
'作业说明_update.html'中，在pilot要求中，也标注了模型横向对比表（示例格式）。你的输出物（excel）中，最好分成两部分。第一部分为类似实例格式的总览表，第二部分为详细情况。你可以适当美化表格。
