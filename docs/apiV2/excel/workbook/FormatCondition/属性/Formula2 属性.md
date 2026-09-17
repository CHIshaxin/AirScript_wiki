---
title: "Formula2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Formula2%20%E5%B1%9E%E6%80%A7.html
---
# Formula2 (属性)

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。

## 示例

```javascript
/*本示例判断如果第一张工作表上单元格区域 E1:E10 的第一个条件格式的公式指定的是“介于 5 和 10 之间”，则对其进行更改。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    if (formatCondition.Operator == xlBetween && formatCondition.Formula1 == "=5" && formatCondition.Formula2 == "=10") {
        formatCondition.Modify(xlCellValue, xlLess, "=10")
    }
}
```

```javascript
/*本示例显示活动工作表上区域 E1:E10 中第一个条件格式的第二部分的表达式。*/
function test() {
    let formatCondition = ActiveSheet.Range("E1:E10").FormatConditions.Item(1)
    console.log(formatCondition.Formula2)
}
```
