# 激励核算 Excel 模板

激励核算输出文件的标准 Sheet 结构、字段定义和公式模式。

## Table of Contents
1. [通用 Sheet 架构](#通用-sheet-架构)
2. [原始数据 Sheet](#原始数据-sheet)
3. [个人明细表 Sheet](#个人明细表-sheet)
4. [部门汇总表 Sheet](#部门汇总表-sheet)
5. [校验报告 Sheet](#校验报告-sheet)
6. [常见激励类型的公式模式](#常见激励类型的公式模式)
7. [条件格式与异常标注](#条件格式与异常标注)

---

## 通用 Sheet 架构

每份激励核算 Excel 文件包含以下 Sheet：

```
┌─────────────────────────────────────────────────────┐
│  原始数据        │ 用户提供的人员信息 + 业绩数据     │
│  (蓝色字体)      │ 作为所有公式的数据源              │
├─────────────────────────────────────────────────────┤
│  个人明细表      │ 每人一行，完整计算链路            │
│  (黑色/绿色字体) │ 中间步骤逐列展开，全部公式        │
├─────────────────────────────────────────────────────┤
│  部门汇总表      │ 按部门/团队分组汇总               │
│  (公式汇总)      │ SUMIF/COUNTIF/AVERAGE             │
├─────────────────────────────────────────────────────┤
│  校验报告        │ 异常值标注、总额校验              │
│  (辅助)          │ 抽样复核记录                      │
└─────────────────────────────────────────────────────┘
```

**Sheet Tab 颜色**：
- 原始数据: `'4472C4'`（蓝色）
- 个人明细表: `'1F4E79'`（深藏青）
- 部门汇总表: `'548235'`（绿色）
- 校验报告: `'BF8F00'`（金色）

---

## 原始数据 Sheet

存放从用户 Excel 中提取的原始数据，**不做任何计算**，仅作为数据源。

### 布局

```
     A          B          C          D          E          F        G
R1:  ─────────── 区域标题：人员信息 ──────────────────────────────────
R2:  工号        姓名       部门       岗位       职级       入职日期  基本工资
R3:  [蓝色]     [蓝色]     [蓝色]     [蓝色]     [蓝色]     [蓝色]   [蓝色]
...

Rn:  ─────────── 区域标题：业绩数据 ──────────────────────────────────
Rn+1: 工号      姓名       销售额     回款额     达成率     ...
Rn+2: [蓝色]    [蓝色]     [蓝色]     [蓝色]     [蓝色]
...
```

### 关键规则
- 所有数据值用 **蓝色字体** 标记（表示是输入值）
- 第一行为区域标题（深藏青底 + 白色文字）
- 第二行为列标题（浅蓝底 + 加粗）
- 数据从第三行开始
- 如果人员信息和业绩数据在同一 Sheet 中，用区域标题分隔
- 如果数据量大或结构差异大，可拆分为"人员信息"和"业绩数据"两个 Sheet

---

## 个人明细表 Sheet

核心计算 Sheet，每人一行，从左到右体现完整计算流程。

### 布局原则

```
     A      B      C      D      E        F          G           H         I
R1:  ───────────── 区域标题：个人激励核算明细 ──────────────────────────────────
R2:  工号   姓名   部门   基本工资  销售额   达成率     绩效系数    基础奖金    最终奖金
                          [绿色↗]  [绿色↗]  [黑色=]   [黑色=]    [黑色=]    [黑色=]
R3:  ...
```

**列的组织顺序**：
1. **标识列**（工号、姓名、部门）— 绿色字体，公式引用原始数据 Sheet
2. **原始数据引用列**（基本工资、销售额等）— 绿色字体，公式引用原始数据 Sheet
3. **中间计算列**（达成率、系数等）— 黑色字体，每列一个计算步骤
4. **最终结果列**（最终奖金）— 黑色加粗字体，中蓝底色

### 公式模式

```python
# 标识列和数据引用列 — 绿色字体，引用原始数据
# 假设原始数据 Sheet 名为 "原始数据"，数据从第3行开始
for i, data_row in enumerate(data_rows):
    row = header_row + 1 + i
    data_r = data_start_row + i  # 原始数据中对应的行号

    # 工号（绿色，跨Sheet引用）
    write_link(ws, row, 1, '原始数据', f'A{data_r}', fmt='@')
    # 姓名
    write_link(ws, row, 2, '原始数据', f'B{data_r}', fmt='@')
    # 部门
    write_link(ws, row, 3, '原始数据', f'C{data_r}', fmt='@')
    # 基本工资
    write_link(ws, row, 4, '原始数据', f'G{data_r}', fmt='#,##0')

    # 中间计算列 — 黑色字体，同Sheet公式
    # 达成率 = 实际销售额 / 目标销售额
    write_formula(ws, row, 6, f'=E{row}/目标!E{data_r}', fmt='0.0%')

    # 最终结果列 — 加粗
    write_formula(ws, row, 9, f'=D{row}*G{row}*H{row}', fmt='#,##0.00', bold=True)
```

### 末尾汇总行

```python
last_data_row = header_row + len(data_rows)
total_row = last_data_row + 1

# 空一行后写汇总
ws.cell(row=total_row, column=1, value='合计').font = BOLD_FONT
# 总金额
write_formula(ws, total_row, 9,
    f'=SUM(I{header_row+1}:I{last_data_row})',
    fmt='#,##0.00', bold=True)
apply_total_border(ws, total_row, 1, 9)

# 人均金额
ws.cell(row=total_row+1, column=1, value='人均').font = BOLD_FONT
write_formula(ws, total_row+1, 9,
    f'=I{total_row}/COUNTA(A{header_row+1}:A{last_data_row})',
    fmt='#,##0.00')
```

---

## 部门汇总表 Sheet

按方案中定义的分组维度（通常是部门）进行汇总统计。

### 布局

```
     A          B        C          D          E          F
R1:  ─────────── 区域标题：部门汇总 ────────────────────────────
R2:  部门        人数     奖金合计    人均奖金    最高奖金    最低奖金
R3:  销售一部    [公式]   [公式]     [公式]     [公式]     [公式]
R4:  销售二部    ...
...
Rn:  ─────────── 总计行（双线下划线）────────────────────────────
```

### 公式模式

```python
# 假设个人明细表 Sheet 名为 "个人明细表"
# 部门在 C 列，最终奖金在 I 列，数据范围第3行到第N行
detail_sheet = '个人明细表'
dept_col = 'C'
amount_col = 'I'
data_range = f'{header_row+1}:{last_data_row}'

for i, dept_name in enumerate(departments):
    row = header_row + 1 + i

    ws.cell(row=row, column=1, value=dept_name).font = BLACK_FONT

    # 人数
    write_formula(ws, row, 2,
        f"=COUNTIF('{detail_sheet}'!{dept_col}{data_range},"
        f'"{dept_name}")',
        fmt='#,##0')

    # 奖金合计
    write_formula(ws, row, 3,
        f"=SUMIF('{detail_sheet}'!{dept_col}{data_range},"
        f'"{dept_name}",\'{detail_sheet}\'!{amount_col}{data_range})',
        fmt='#,##0.00')

    # 人均奖金
    write_formula(ws, row, 4,
        f'=IF(B{row}=0,"",C{row}/B{row})',
        fmt='#,##0.00')

    # 最高奖金（使用 MAXIFS，Excel 2019+）
    write_formula(ws, row, 5,
        f"=MAXIFS('{detail_sheet}'!{amount_col}{data_range},"
        f"'{detail_sheet}'!{dept_col}{data_range},"
        f'"{dept_name}")',
        fmt='#,##0.00')

    # 最低奖金
    write_formula(ws, row, 6,
        f"=MINIFS('{detail_sheet}'!{amount_col}{data_range},"
        f"'{detail_sheet}'!{dept_col}{data_range},"
        f'"{dept_name}")',
        fmt='#,##0.00')

# 总计行
total_row = header_row + 1 + len(departments)
ws.cell(row=total_row, column=1, value='总计').font = BOLD_FONT
write_formula(ws, total_row, 2, f'=SUM(B{header_row+1}:B{total_row-1})', fmt='#,##0')
write_formula(ws, total_row, 3, f'=SUM(C{header_row+1}:C{total_row-1})', fmt='#,##0.00')
write_formula(ws, total_row, 4, f'=IF(B{total_row}=0,"",C{total_row}/B{total_row})', fmt='#,##0.00')
apply_total_border(ws, total_row, 1, 6)
```

---

## 校验报告 Sheet

用于质量检查和结果验证。

### 布局

```
     A                    B              C
R1:  ─────── 区域标题：核算校验报告 ────────────
R2:  校验项               结果           状态
R3:  参与核算人数          [公式]         ✅
R4:  总奖金金额            [公式]         ✅
R5:  人均奖金              [公式]         ✅
R6:  奖金为0人数           [公式]         ⚠️
R7:  奖金为负数人数        [公式]         🔴
R8:  最高奖金              [公式]
R9:  最低奖金（>0）        [公式]
R10: 最高/最低比值          [公式]

     ─────── 区域标题：异常明细 ──────────────
R12: 工号    姓名    部门    异常类型    奖金金额
...
```

### 公式模式

```python
# 校验项公式
checks = [
    ('参与核算人数', f"=COUNTA('{detail_sheet}'!A{h+1}:A{last})", '#,##0'),
    ('总奖金金额', f"=SUM('{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last})", '#,##0.00'),
    ('人均奖金', f'=B{row_total}/B{row_count}', '#,##0.00'),
    ('奖金为0人数', f"=COUNTIF('{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last},0)", '#,##0'),
    ('奖金为负数人数', f"=COUNTIF('{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last},\"<0\")", '#,##0'),
    ('最高奖金', f"=MAX('{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last})", '#,##0.00'),
    ('最低奖金（>0）', f"=MINIFS('{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last},"
                       f"'{detail_sheet}'!{amt_col}{h+1}:{amt_col}{last},\">0\")", '#,##0.00'),
]

# 状态列：条件判断
# 奖金为负数 > 0 → 🔴；奖金为0人数 > 0 → ⚠️；其他 → ✅
```

### 异常明细区域

```python
# 筛选并列出异常记录（奖金为0、为负数、金额超出合理范围等）
# 使用 Python 逻辑筛选后写入，标记异常类型
anomaly_types = {
    'zero': '奖金为零',
    'negative': '奖金为负数',
    'extreme_high': '金额异常偏高',
    'extreme_low': '金额异常偏低',
}

# 异常行用 ERROR_FILL（浅红底色）标记
for row_data in anomalies:
    for col_idx in range(1, len(row_data) + 1):
        cell = ws.cell(row=current_row, column=col_idx, value=row_data[col_idx-1])
        cell.fill = ERROR_FILL
```

---

## 常见激励类型的公式模式

### 类型 1：业绩达成率 × 系数

```
最终奖金 = 基本工资 × 奖金基数比例 × 达成率系数

其中达成率系数为阶梯制：
  达成率 < 60%  → 系数 0
  60% ≤ 达成率 < 80%  → 系数 0.6
  80% ≤ 达成率 < 100% → 系数 0.8 + (达成率-80%)/20% × 0.2
  100% ≤ 达成率 < 120% → 系数 1.0
  达成率 ≥ 120% → 系数 1.2（封顶）
```

**Excel 公式示例**（假设达成率在 F 列）：
```
=IF(F3<0.6, 0, IF(F3<0.8, 0.6, IF(F3<1, 0.8+(F3-0.8)/0.2*0.2, IF(F3<1.2, 1, 1.2))))
```

> 建议将阶梯系数表放在单独的"参数"区域，用 VLOOKUP/INDEX+MATCH 引用，避免公式过长。

### 类型 2：项目奖金分配

```
个人奖金 = 项目总奖金包 × 个人权重 / 权重合计

其中：
  个人权重 = 角色权重 × 参与度 × 绩效评级系数
```

**计算列拆分**：
```
D: 角色权重        [绿色，引用参数表]
E: 参与度          [绿色，引用原始数据]
F: 绩效评级系数    [绿色，引用原始数据]
G: 个人权重        = D × E × F
H: 权重占比        = G / SUM(G列)
I: 最终奖金        = 奖金总包 × H
```

### 类型 3：销售提成

```
提成金额 = 回款金额 × 提成比例

阶梯提成：
  回款 ≤ 50万  → 3%
  50万 < 回款 ≤ 100万 → 5%
  回款 > 100万 → 8%（超额部分）
```

**建议处理方式**：将阶梯区间和比例放在参数区域，用分段计算公式：
```
=MIN(E3, 500000)*0.03 + MAX(MIN(E3, 1000000)-500000, 0)*0.05 + MAX(E3-1000000, 0)*0.08
```

---

## 条件格式与异常标注

```python
def mark_anomaly(ws, row, col, value, anomaly_type, fmt='#,##0.00'):
    """标记异常值（红色字体 + 浅红底色）。"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = RED_FONT
    cell.fill = ERROR_FILL
    cell.number_format = fmt
    cell.comment = Comment(f'异常: {anomaly_type}', '校验')

def mark_pass(ws, row, col, value, fmt='#,##0.00'):
    """标记校验通过（浅绿底色）。"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = PASS_FILL
    cell.number_format = fmt

def highlight_output_row(ws, row, col_start, col_end):
    """高亮最终结果列（中蓝底色）。"""
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = OUTPUT_FILL
```
