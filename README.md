# 激励核算执行工具 (Incentive Calculation Skill)

> CodeBuddy Skill — 根据已确定的激励方案，结合业绩数据和人员信息，计算每个人的实际激励金额，并输出咨询级品质的 Excel 核算文件。

## 功能概述

本 Skill 专注于**激励方案的计算执行**，不做方案设计，只做核算。

### 三阶段工作流

```
Phase 1 理解方案 → Phase 2 数据校验 → Phase 3 执行计算
    ↓ 用户确认        ↓ 用户确认          ↓ 输出结果
```

### 覆盖激励类型

- 业绩奖金（年度 / 季度 / 月度）
- 项目奖金
- 销售提成
- 其他自定义激励方案

### Excel 输出标准

- **咨询级格式**：蓝色=输入值、黑色=公式、绿色=跨Sheet引用
- **全公式可追溯**：所有计算保留 Excel 公式，不写死数值
- **标准 Sheet 结构**：原始数据 → 个人明细表 → 部门汇总表 → 校验报告

## 文件结构

```
├── SKILL.md                           # 主文件：完整的工作流和规范
└── references/
    ├── data_validation.md             # 数据校验规则参考
    ├── incentive-templates.md         # 激励核算 Excel 模板定义
    └── openpyxl-patterns.md           # openpyxl 代码模式和可复用函数
```

## 使用方式

本 Skill 为 [CodeBuddy](https://www.codebuddy.ai) 的自定义 Skill，安装后可通过以下关键词触发：

> 激励核算、奖金计算、奖金核算、激励计算、业绩奖金计算、项目奖金核算、薪酬核算、奖金发放、激励测算、核算明细、奖金明细、激励发放表

## 安装

将本仓库克隆到 CodeBuddy Skills 目录：

```bash
git clone https://github.com/sean-li-git/incentive-calculation-skill.git ~/.codebuddy/skills/incentive-calculation
```

## License

MIT
