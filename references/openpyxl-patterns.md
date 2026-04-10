# openpyxl 代码模式

激励核算 Excel 文件的代码模式和可复用函数。

## Table of Contents
1. [工作簿初始化](#工作簿初始化)
2. [区域标题](#区域标题)
3. [输入单元格（蓝色字体）](#输入单元格)
4. [公式单元格（黑色字体）](#公式单元格)
5. [跨Sheet引用（绿色字体）](#跨sheet引用)
6. [数字格式](#数字格式)
7. [边框与汇总行](#边框与汇总行)
8. [数据验证与下拉框](#数据验证)
9. [图表](#图表)
10. [最终美化](#最终美化)

---

## 工作簿初始化

```python
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    numbers, NamedStyle
)
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

wb = Workbook()

# === 字体颜色常量 ===
BLUE_FONT = Font(name='Calibri', size=10, color='0000FF')       # 输入值/原始数据
BLACK_FONT = Font(name='Calibri', size=10, color='000000')       # 同Sheet公式
GREEN_FONT = Font(name='Calibri', size=10, color='008000')       # 跨Sheet引用
RED_FONT = Font(name='Calibri', size=10, color='FF0000')         # 异常标记

HEADER_FONT = Font(name='Calibri', size=11, color='FFFFFF', bold=True)  # 区域标题
BOLD_FONT = Font(name='Calibri', size=10, bold=True)
BOLD_BLUE = Font(name='Calibri', size=10, color='0000FF', bold=True)
ITALIC_FONT = Font(name='Calibri', size=10, italic=True)

# === 填充色 ===
NAVY_FILL = PatternFill('solid', fgColor='1F4E79')        # 区域标题背景
LIGHT_BLUE_FILL = PatternFill('solid', fgColor='D9E1F2')  # 列标题
YELLOW_FILL = PatternFill('solid', fgColor='FFF2CC')       # 关键输入高亮
OUTPUT_FILL = PatternFill('solid', fgColor='BDD7EE')       # 输出/汇总行
GRAY_FILL = PatternFill('solid', fgColor='F2F2F2')         # 隔行底色
WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')
ERROR_FILL = PatternFill('solid', fgColor='FFC7CE')        # 异常值高亮（浅红）
PASS_FILL = PatternFill('solid', fgColor='C6EFCE')         # 校验通过（浅绿）

# === 边框 ===
THIN_BOTTOM = Border(bottom=Side(style='thin'))
DOUBLE_BOTTOM = Border(bottom=Side(style='double'))
THIN_TOP = Border(top=Side(style='thin'))
BOX_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# === 对齐 ===
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')
INDENT_1 = Alignment(horizontal='left', vertical='center', indent=1)
INDENT_2 = Alignment(horizontal='left', vertical='center', indent=2)
WRAP_TEXT = Alignment(horizontal='left', vertical='top', wrap_text=True)
```

## 区域标题

```python
def write_section_header(ws, row, col_start, col_end, title):
    """深藏青标题行，跨多列合并。"""
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=title)
    cell.font = HEADER_FONT
    cell.fill = NAVY_FILL
    cell.alignment = LEFT_ALIGN
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = NAVY_FILL

def write_sub_header(ws, row, col_start, col_end, title):
    """浅蓝子标题行。"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
    ws.cell(row=row, column=col_start, value=title)

def write_column_headers(ws, row, labels, start_col=1):
    """写入列标题行（字段名）。"""
    for i, label in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + i, value=label)
        cell.font = BOLD_FONT
        cell.fill = LIGHT_BLUE_FILL
        cell.alignment = CENTER_ALIGN
```

## 输入单元格

```python
def write_input(ws, row, col, value, fmt='#,##0', comment_text=None):
    """写入蓝色字体的输入值（原始数据）。"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = BLUE_FONT
    cell.number_format = fmt
    if comment_text:
        cell.comment = Comment(comment_text, 'Model')

def write_input_row(ws, row, values, start_col=1, fmt='#,##0'):
    """写入一整行蓝色字体的输入值。"""
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=start_col + i, value=v)
        cell.font = BLUE_FONT
        cell.number_format = fmt

def write_highlighted_input(ws, row, col, value, fmt='#,##0'):
    """写入黄色高亮的关键输入值。"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = BOLD_BLUE
    cell.number_format = fmt
    cell.fill = YELLOW_FILL
```

## 公式单元格

```python
def write_formula(ws, row, col, formula, fmt='#,##0.00', bold=False):
    """写入黑色字体的公式。"""
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = BOLD_FONT if bold else BLACK_FONT
    cell.number_format = fmt

def write_formula_row(ws, row, formulas, start_col=1, fmt='#,##0.00', bold=False):
    """写入一整行公式。"""
    for i, f in enumerate(formulas):
        cell = ws.cell(row=row, column=start_col + i, value=f)
        cell.font = BOLD_FONT if bold else BLACK_FONT
        cell.number_format = fmt

def write_total_row(ws, row, label, label_col, formula_col, formula,
                    fmt='#,##0.00', border_style='single'):
    """写入加粗汇总行。"""
    lbl = ws.cell(row=row, column=label_col, value=label)
    lbl.font = BOLD_FONT
    cell = ws.cell(row=row, column=formula_col, value=formula)
    cell.font = BOLD_FONT
    cell.number_format = fmt
    border = THIN_BOTTOM if border_style == 'single' else DOUBLE_BOTTOM
    cell.border = border
```

## 跨Sheet引用

```python
def write_link(ws, row, col, sheet_name, cell_ref, fmt='#,##0'):
    """写入绿色字体的跨Sheet引用。"""
    formula = f"='{sheet_name}'!{cell_ref}"
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = GREEN_FONT
    cell.number_format = fmt
```

## 数字格式

```python
# 激励核算常用格式
FORMATS = {
    'currency':    '#,##0.00',          # 1,234.56（金额，精确到分）
    'currency_0':  '#,##0',             # 1,234（金额，整数）
    'currency_w':  '#,##0.00"万"',      # 1,234.56万
    'pct':         '0.0%',              # 25.3%
    'pct_0':       '0%',                # 25%
    'pct_2':       '0.00%',             # 25.30%
    'integer':     '#,##0',             # 1,234
    'coefficient': '0.00',              # 1.25（系数）
    'negative':    '#,##0.00;(#,##0.00)',  # 负数用括号
    'text':        '@',                 # 强制文本
}
```

## 边框与汇总行

```python
def apply_subtotal_border(ws, row, col_start, col_end):
    """小计行上方单线。"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        existing = cell.border
        cell.border = Border(
            top=Side(style='thin'),
            left=existing.left, right=existing.right, bottom=existing.bottom
        )

def apply_total_border(ws, row, col_start, col_end):
    """总计行下方双线。"""
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        existing = cell.border
        cell.border = Border(
            bottom=Side(style='double'),
            left=existing.left, right=existing.right, top=existing.top
        )

def apply_box_borders(ws, min_row, max_row, min_col, max_col):
    """为数据区域添加网格线框。"""
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = BOX_BORDER
```

## 数据验证

```python
def add_dropdown(ws, cell_ref, options_list):
    """添加下拉选择框。"""
    dv = DataValidation(
        type='list',
        formula1=f'"{",".join(options_list)}"',
        allow_blank=False
    )
    ws.add_data_validation(dv)
    dv.add(ws[cell_ref])
```

## 图表

```python
def add_bar_chart(ws, title, data_range, cat_range, position,
                  width=15, height=10):
    """添加柱状图。"""
    chart = BarChart()
    chart.type = 'col'
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.add_data(data_range, titles_from_data=True)
    chart.set_categories(cat_range)
    chart.legend.position = 'b'
    ws.add_chart(chart, position)
    return chart

def add_line_chart(ws, title, data_range, cat_range, position,
                   width=15, height=10):
    """添加折线图。"""
    chart = LineChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height
    chart.add_data(data_range, titles_from_data=True)
    chart.set_categories(cat_range)
    chart.legend.position = 'b'
    ws.add_chart(chart, position)
    return chart
```

## 最终美化

```python
def polish_workbook(wb):
    """对所有 Sheet 执行最终格式化。"""
    for ws in wb.worksheets:
        # 自动列宽
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if cell.value:
                        # 中文字符按2倍宽度计算
                        val = str(cell.value)
                        length = sum(2 if ord(c) > 127 else 1 for c in val)
                        max_len = max(max_len, length)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        # 冻结窗格（固定表头）
        if ws.max_row > 1:
            ws.freeze_panes = 'A2'

        # 打印设置
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

def save_workbook(wb, filename):
    """保存并打印确认信息。"""
    wb.save(filename)
    print(f'✅ 文件已保存: {filename}')
```
