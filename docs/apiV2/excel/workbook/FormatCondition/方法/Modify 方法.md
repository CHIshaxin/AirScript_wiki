---
title: "Modify 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E6%96%B9%E6%B3%95/Modify%20%E6%96%B9%E6%B3%95.html
---
# Modify (方法)

更改现有条件格式。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlFormatConditionType | 必填 | 指定条件格式是基于单元格值还是基于表达式。 |
| Operator | any | 可选 | 一个 XlFormatConditionOperator 值，它代表条件格式运算符。如果 Type 设为 xlExpression，则忽略该参数。 |
| Formula1 | any | 可选 | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |
| Formula2 | any | 可选 | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |

## 示例

```javascript
/*本示例修改第一张工作表上单元格区域 E1:E10 的现有条件格式。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    formatCondition.Modify(xlCellValue, xlLess, "=$B$1")
}
```

```javascript
/*本示例修改活动工作表上单元格区域 A1:A10 的现有条件格式。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Modify(xlCellValue, xlBetween, "=3", "=9")
}
```
