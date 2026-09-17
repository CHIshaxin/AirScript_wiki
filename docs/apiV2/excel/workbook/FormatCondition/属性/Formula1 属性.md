---
title: "Formula1 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Formula1%20%E5%B1%9E%E6%80%A7.html
---
# Formula1 (属性)

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String** 型，只读。

## 示例

```javascript
/*本示例判断如果第一张工作表单元格区域 E1:E10 的第一个条件格式的公式指定的是“小于 5”，则此示例将对其进行更改。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    if (formatCondition.Operator == xlLess && formatCondition.Formula1 == "=5") {
        formatCondition.Modify(xlCellValue, xlLess, "=10")
    }
}
```

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式的第一部分的表达式。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.Formula1)
}
```
