# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a financial fraud classification assignment for a university course ("智能财务前沿"). The task is to annotate Chinese regulatory violation records from `STK_Violation_Main_第2组.xlsx` with 6 structured fields indicating annual report relevance, financial impact, and third-party collusion.

**Core workflow**: Rule-based NLP annotation using keyword/pattern matching on the `Activity` field (Chinese regulatory text describing violations).

## Key Commands

```bash
# Run full annotation pipeline
python "全量数据标注/annotate_v2.py"
python "全量数据标注/corrections.py"

# Run pilot phase annotations (5% sample, 3 models)
python Mimo/annotate_v2.py
python DeepSeek/annotate.py
python MiniMax/annotate.py

# Generate validation comparison file
python gen_validation_final.py
```

**Important**: All Python scripts use `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` at the top to handle Chinese characters in Windows terminal.

## Architecture

### Data Pipeline

```
STK_Violation_Main_第2组.xlsx (3724 rows, 23 cols)
    ↓
annotate_v2.py → 3721 annotated rows (skips 3 empty Activity rows)
    ↓
corrections.py → manual fixes for edge cases
    ↓
Rebuild script (inline) → STK_labeled_夏思远_2023110537_G02.xlsx (29 cols)
```

### Output Format

The output file preserves all 23 original columns and appends 6 annotation columns (Col 24-29):

| Column | Field | Values |
|--------|-------|--------|
| 24 | ann_related | 0/1 |
| 25 | ann_year | JSON array string or "null" |
| 26 | ann_fin_flag | 0/1/"null" |
| 27 | ann_fin_info | JSON string or "null" |
| 28 | third_party_flag | 0/1/"null" |
| 29 | third_party_list | JSON string or "null" |

**Critical**: null values must be the string `"null"`, not Python None. JSON fields must be valid JSON strings.

### Annotation Logic (annotate_v2.py)

Three-task pipeline with conditional dependencies:

1. **Task 1** (`is_annual_report_related`): Determines if violation affects annual reports
   - Excludes: window trading, inquiry letter responses, insider trading, governance issues
   - Requires "年度报告" or "年报" keyword in Activity text

2. **Task 2** (`is_financial_info` / `build_ann_fin_info`): Only if Task 1=1
   - Distinguishes financial vs non-financial violations
   - Maps to 6 accounting elements: 资产/负债/所有者权益/收入/费用/利润
   - Year-element pairs must correspond

3. **Task 3** (`identify_third_party`): Only if Task 1=1 AND Task 2=1
   - Identifies colluding third parties (customers, suppliers, etc.)
   - Excludes: subsidiaries, controlling shareholders, actual controllers
   - Type inference uses Activity context, not just entity name

### Data Alignment

The source file has 3 header/unit rows (rows 1-3) and 3 empty Activity rows (rows 28, 1325, 2060). When rebuilding the output file, empty Activity rows must get "null" annotations — do NOT shift annotation data to fill gaps.

### Key Files

| File | Purpose |
|------|---------|
| `全量数据标注/annotate_v2.py` | Main annotation script (Task 1/2/3 logic) |
| `全量数据标注/corrections.py` | Manual corrections for edge cases |
| `全量数据标注/第三方配合造假判断经验手册.md` | Domain knowledge for third-party collusion detection |
| `Mimo/`, `DeepSeek/`, `MiniMax/` | Pilot phase model annotations (5% sample) |
| `validation_夏思远_2023110537_G02.xlsx` | Model comparison F1 scores |
| `作业说明_update.html` | Assignment specification (judgment criteria) |

## Assignment Deliverables

1. `STK_labeled_夏思远_2023110537_G02.xlsx` — Full 3721-row annotated output
2. `report_夏思远_2023110537_G02.pdf` — Workflow, prompts, model comparison, error analysis
3. `validation_夏思远_2023110537_G02.xlsx` — 5% sample with human vs model annotations + F1 scores
