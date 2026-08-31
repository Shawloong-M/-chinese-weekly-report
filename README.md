# 中文周报/月报 PPTX 智能生成器 (Chinese Weekly Report Generator)

一键生成专业级中文周报/月报 PPTX，专为制造业、运营、数字化岗位设计。输入 JSON 数据，输出 6 页完整 PPTX。

## 功能特性

- 📊 **6 页完整结构**：封面 → 目录 → 本周概述 → 核心KPI → 重点工作与问题对策 → 下周计划
- 📈 **KPI 看板**：左右分栏，左侧趋势图 + 右侧指标卡 + 达标路径
- 📋 **10 列任务进度表**：含序号、项目名称、内容、进度、责任人、起止时间、状态、备注
- 📅 **甘特图**：可编辑的条形图，每项任务可拖拽调整
- 🏭 **工厂风格封面**：深蓝底色 + 建筑装饰 + LOGO 占位
- 🔄 **双输出**：PPTX（可编辑）+ HTML（直接预览）
- 💰 **数据驱动**：所有内容来自 JSON 配置文件，改数据不改模板

## 快速开始

### 1. 安装依赖

```bash
pip install python-pptx
```

### 2. 准备数据

复制 `sample_data.json`，按实际数据修改。

### 3. 生成 PPTX

```bash
python scripts/chinese_weekly_report.py your_data.json output.pptx
```

### 4. 打开查看

双击生成的 `output.pptx` 用 PowerPoint 打开，所有元素均可编辑。

## 数据格式

参考 `sample_data.json`，核心字段说明：

| 字段                | 说明                         | 必填 |
| ------------------- | ---------------------------- | ---- |
| `report_period`     | 报告周期（如"2026年第35周"） | 是   |
| `department`        | 部门名称                     | 是   |
| `overview_table`    | 任务进度表（5-6行）          | 是   |
| `overview_gantt`    | 甘特图任务列表               | 是   |
| `kpi_right_modules` | KPI 指标卡（最多4个）        | 推荐 |
| `kpi_chart`         | 趋势图配置                   | 推荐 |
| `key_works_table`   | 重点工作表                   | 是   |
| `problems_table`    | 问题对策表                   | 推荐 |
| `next_plan_table`   | 下周计划表                   | 是   |
| `next_plan_gantt`   | 下周甘特图                   | 推荐 |

## 输出结构

| 页码 | 内容               | 说明                                     |
| ---- | ------------------ | ---------------------------------------- |
| 1    | 封面               | 标题 + LOGO + 部门/报告人/日期           |
| 2    | 目录               | 四章导航                                 |
| 3    | 本周概述           | 总结 + 10列任务表 + 甘特图               |
| 4    | 核心KPI            | 55/45 分栏：左侧描述+图表，右侧指标+路径 |
| 5    | 重点工作与问题对策 | 合并页，各带总结                         |
| 6    | 下周计划           | 计划表 + 甘特图                          |

## 截图

<!-- 建议放一张生成后的 PPTX 截图 -->

## 技术栈

- Python 3.10+
- python-pptx
- Chart.js (HTML 输出模式)

## 适用场景

- 制造业周报/月报
- 数字化部门工作报告
- 生产运营 KPI 汇报
- 项目进度汇报

## 支持

如果这个项目帮到了你，欢迎请我喝杯咖啡 ☕

[https://afdian.com/a/Shawloong](https://afdian.com/a/Shawloong)

## 许可证

MIT

## 作者

[mxl_t](https://github.com/mxl_t) - 数字化从业者，Hermes Agent 技能开发者
