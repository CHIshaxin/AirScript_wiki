---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

添加新的条件格式。返回一个 **[FormatCondition](./../../FormatCondition/FormatCondition 对象.html)** 对象，它代表新的条件格式。

## 说明

对单个区域定义的条件格式不能超过三个。使用 **Modify** 方法可修改现有的条件格式，使用 **Delete** 方法可在添加新条件格式前删除现有的格式。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlFormatConditionType | 必填 | 指定条件格式是基于单元格值还是基于表达式。 |
| Operator | any | 可选 | 条件格式运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。如果 Type 为 xlExpression，则忽略 Operator 参数。 |
| Formula1 | any | 可选 | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |
| Formula2 | any | 可选 | 当 Operator 为 xlBetween 或 xlNotBetween 时，它是与条件格式第二部分相关联的值或表达式（否则忽略该参数）。可为常量值、字符串值、单元格引用或公式。 |

## 示例

```javascript
/*本示例向第一张工作表上单元格区域 E1:E10 中添加条件格式*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Add(xlCellValue, xlGreater, "=$A$1")
    let boders = formatCondition.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThin
    boders.ColorIndex = 6
    let font = formatCondition.Font
    font.Bold = true
    font.ColorIndex = 3
}
```

```javascript
/*本示例在活动工作表上单元格区域 B1:B10 上新增条件格式，并设置该条件格式的内部颜色。*/
function test() {
    let formatCondition = ActiveSheet.Range("B1:B10").FormatConditions.Add(xlCellValue, xlBetween, "=3", "=9")
    formatCondition.Interior.ColorIndex = 3
}
```
