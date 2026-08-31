---
name: chinese-weekly-report
description: Generate Chinese weekly/monthly reports with KPI cards.
version: 2.2.0
author: mxl_t (mxl_t), Hermes Agent
license: MIT
platforms: [windows, macos, linux]
metadata:
  hermes:
    tags: [chinese, report, weekly, monthly, kpi, pptx, data-driven, manufacturing]
    category: productivity
    related_skills: [powerpoint, data-driven-presentation, xlsx]
---

# Chinese Weekly Report Skill

生成专业的中文周报/月报 PPTX，包含工厂封面（带LOGO占位）、目录页、KPI 看板（左右分栏）、甘特图、任务进度表等。专为制造业、运营、数字化岗位设计。

## When to Use

- The user asks to **generate a weekly/monthly report in Chinese** from raw data.
- The user needs a **professional manufacturing-style report** with cover, TOC, Gantt charts, and KPI dashboard.
- The user is in **manufacturing, factory, chip, or industrial digitalization** roles.
- Do NOT use for: English-only reports, creative presentations, simple text reports.

## Prerequisites

- `python-pptx` installed: `pip install python-pptx`
- Python 3.10+ with python-pptx installed in the active interpreter
- Source data: JSON spec file (see below)

## Quick Reference

```
python scripts/chinese_weekly_report.py data.json report.pptx
```

## Data JSON Structure

```json
{
  "report_type": "weekly",
  "report_period": "2026年第35周",
  "department": "数字化推进部",
  "author": "张三",
  "date": "2026-08-31",
  "logo_text": "公司LOGO",

  "overview_summary": "本周重点推进MES系统数据对接项目...",

  "overview_table": [
    {
      "seq": 1, "project": "MES系统数据对接", "content": "PLC数据采集接口开发",
      "weekly_progress": "完成3条产线对接", "owner": "张三",
      "start": "2026-08-01", "end": "2026-09-15",
      "progress": 100, "status": "已完成", "remark": ""
    }
  ],

  "overview_gantt": {
    "start_date": "2026-08-01",
    "end_date": "2026-09-30",
    "tasks": [
      {"name": "MES系统对接", "start": "2026-08-01", "end": "2026-09-15", "color": "#3A7BD5"},
      {"name": "看板设计", "start": "2026-08-15", "end": "2026-09-10", "color": "#27AE60"}
    ]
  },

  "kpi_summary": "本周6项核心KPI中，4项达标，2项未达标。",
  "kpi_left_title": "指标达成分析",
  "kpi_left_desc": "产量达成率98.5%，连续3周保持在95%以上；产品良率97.3%创历史新高；异常响应时间12min，较目标值差2min，需重点关注。",
  "kpi_chart": {
    "type": "line",
    "title": "近8周良率趋势",
    "categories": ["W28","W29","W30","W31","W32","W33","W34","W35"],
    "series": [
      {"name": "目标良率", "values": [95, 95, 95, 95, 95, 95, 95, 95]},
      {"name": "实际良率", "values": [91.2, 92.5, 93.1, 94.0, 94.8, 95.2, 96.1, 97.3]}
    ]
  },
  "kpi_right_modules": [
    {
      "name": "产量达成率",
      "value": "98.5%",
      "target": "≥95%",
      "trend": "up",
      "good": true,
      "path": "优化排产算法→提升设备利用率→减少换型时间"
    }
  ],
  "kpi_right_path_title": "达标路径与逻辑",

  "key_works_summary": "本周共推进4项重点工作...",
  "key_works_table": [
    {
      "seq": 1, "project": "MES系统数据接口开发", "content": "PLC数据采集接口开发",
      "weekly_progress": "完成3条产线对接", "owner": "张三",
      "start": "2026-08-01", "end": "2026-09-15",
      "progress": 100, "status": "已完成", "remark": ""
    }
  ],
  "problems_summary": "本周存在2个突出问题...",
  "problems_table": [
    {
      "issue": "部分老旧设备不支持标准通讯协议",
      "impact": "影响数据采集覆盖率",
      "solution": "采购协议转换网关，预计成本2万元",
      "deadline": "2026-09-15",
      "owner": "李四",
      "status": "进行中"
    }
  ],

  "next_plan_summary": "下周计划推进6项工作...",
  "next_plan_table": [
    {
      "seq": 1, "project": "完成剩余产线对接", "content": "PLC数据采集",
      "next_plan": "完成2条产线安装调试", "owner": "张三",
      "start": "2026-09-01", "end": "2026-09-07",
      "progress": 0, "status": "计划中", "remark": ""
    }
  ],
  "next_plan_gantt": {
    "start_date": "2026-09-01",
    "end_date": "2026-09-30",
    "tasks": [
      {"name": "产线对接", "start": "2026-09-01", "end": "2026-09-07", "color": "#3A7BD5"}
    ]
  }
}
```

## Slide Structure

| Slide | Content |
|-------|---------|
| 1 | 封面 — 工厂背景 + LOGO占位(左上) + 标题/周期/部门 |
| 2 | 目录 — 章节列表 |
| 3 | 本周概述 — 总结(顶) + 10列任务进度表(中) + 甘特图(底) |
| 4 | 核心KPI — 55/45左右分栏：左侧描述+趋势图，右侧指标卡+达标路径 |
| 5 | 重点工作与问题对策 — 合并页，各带1行总结 |
| 6 | 下周计划 — 10列计划表 + 甘特图 |

## Tables

10列任务进度表：序号、项目名称、项目内容、本周进度、责任人、开始时间、结束时间、进度、状态、备注。
- 表头颜色：中蓝 `#2B5C9A`，字体 9pt 加粗
- 数据行：字体 8pt，交替行颜色 `#F0F4F8` / 白色
- 6列问题表：问题描述、影响、解决方案、完成时限、责任人、状态

## Gantt Chart

- 任务名称列宽 1.5 英寸
- 条形高度为行高的 60%，带边框
- 交替行浅灰背景
- 每个任务条支持点击拖拽编辑（PowerPoint 原生形状）

## KPI Page Layout

- 55/45 左右分栏（左侧 55% 给图表和描述，右侧 45% 给指标卡）
- 左侧：总结 → 描述框 → 折线/柱状图
- 右侧：4个指标卡（名称 + 值 + 趋势箭头 + 目标）→ 达标路径列表

## Verification

1. Run `python scripts/chinese_weekly_report.py data.json out.pptx`
2. Verify output JSON: `{"output": "out.pptx", "slides": 6}`
3. Open in PowerPoint and verify:
   - Cover has LOGO placeholder (top-left)
   - Each page has LOGO in top-left
   - Tables render correctly with all 10 columns
   - Gantt charts are editable shapes with alternating row backgrounds
   - KPI page uses 55/45 left-right split
   - Key Works + Problems are merged on one page