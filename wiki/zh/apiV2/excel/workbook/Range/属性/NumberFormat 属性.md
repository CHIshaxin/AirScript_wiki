---
title: "NumberFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CAD80COo
---

# NumberFormat (属性)

返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

如果指定区域中的所有单元格包含不同的数字格式，则此属性返回 **Null**。

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*以下这些示例分别对 Sheet1 中的 A17 单元格、第一行和 C 列的数字格式进行设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A17").NumberFormat = "General"
    Application.Worksheets.Item("Sheet1").Rows.Item(1).NumberFormat = "hh:mm:ss"
    Application.Worksheets.Item("Sheet1").Columns.Item("C").NumberFormat = "$#,##0.00_);Red"
}
```

```javascript
/*本示例将 A1:B3 单元格区域的数字格式设置为 0.00%*/
function test() {
    Range("A1:B3").NumberFormat = "0.00%"
}
```
