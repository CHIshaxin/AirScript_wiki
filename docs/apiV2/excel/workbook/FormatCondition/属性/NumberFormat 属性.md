---
title: "NumberFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormat (属性)

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。**Variant** 型，可读写。

## 说明

数字格式是使用 **“单元格格式”** 对话框的 **“数字”** 选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 `"General"` 或者创建自定义数字格式。

## 示例

```javascript
/*本示例显示第一张工作表上单元格区域 A1:A10 的第一个条件格式所应用于的数字格式。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.NumberFormat)
}
```

```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第一个条件格式所应用于的数字格式。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.NumberFormat = "0%"
}
```
