---
title: "UseStandardWidth 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/UseStandardWidth%20%E5%B1%9E%E6%80%A7.html
---
# UseStandardWidth (属性)

如果 **[Range](./../../Range/Range 对象.html)** 对象的列宽等于工作表的标准列宽，则该属性的值为 **True**。如果区域包含不止一列并且不是所有的列都等宽，则返回 **null**。**Variant** 类型，可读写。

## 示例

```javascript
/*本示例将工作表 Sheet1 上的 A 列的列宽设置为标准列宽。*/
function test() {
    Worksheets.Item("Sheet1").Columns.Item("A").UseStandardWidth = true
}
```

```javascript
/*本示例将 B1 单元格所在列的列宽设置为 40，并显示 B1 单元格和 A1:D1 单元格区域的列宽是否等于标准列宽*/
function test() {
    Range("B1").Columns.ColumnWidth = 40
    console.log(Range("B1").UseStandardWidth)
    console.log(Range("A1:D1").UseStandardWidth)
}
```
