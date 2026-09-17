---
title: "Operator 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Operator%20%E5%B1%9E%E6%80%A7.html
---
# Operator (属性)

返回一个 **Long** 值，它代表条件格式的操作符。

## 示例

```javascript
/*本示例判断如果第一张工作表单元格区域 E1:E10 的第一个条件格式的公式指定的是“小于 5”，则此示例将对其进行更改。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    if (formatCondition.Operator == xlLess && formatCondition.Formula1 == "=5") {
        formatCondition.Modify(xlCellValue, xlBetween, "=5", "=15")
    }
}
```

```javascript
/*本示例显示活动工作表单元格区域 E1:E10 的第一个条件格式的操作符是否为xlLess。*/
function test() {
    let formatCondition = ActiveSheet.Range("E1:E10").FormatConditions.Item(1)
    console.log(formatCondition.Operator == xlLess)
}
```
