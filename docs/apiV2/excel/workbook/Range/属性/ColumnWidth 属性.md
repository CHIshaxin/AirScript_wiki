---
title: "ColumnWidth 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ColumnWidth%20%E5%B1%9E%E6%80%A7.html
---
# ColumnWidth (属性)

返回或设置指定区域中所有列的列宽。**Variant** 类型，可读写。

## 说明

一个列宽单位等于“常规”样式中一个字符的宽度。对于比例字体，则使用字符 0（零）的宽度。

使用 **Width** 属性可返回以磅为单位的列宽。

如果区域中所有列的列宽都相等，**ColumnWidth** 属性返回该宽度值。如果区域中的列宽不等，该属性返回 **null**。

## 示例

```javascript
/*本示例使工作表 Sheet1 上 A 列的列宽加倍。*/
function test() {
    let itm = Application.Worksheets.Item("Sheet1").Columns.Item("A")
    itm.ColumnWidth = itm.ColumnWidth * 2
}
```

```javascript
/*本示例将工作表 Sheet1 上 A1 单元格的列宽设置为5。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").ColumnWidth = 5
}
```
