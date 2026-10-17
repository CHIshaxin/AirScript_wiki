---
title: "Cells (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Cells%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cdha91VQ
---

# Cells (属性)

返回一个 Range对象，它代表工作表中的所有单元格（不仅仅是当前使用的单元格）。

## 说明

因为**Item**属性是 Range 对象的默认属性，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 Range 对象，它代表活动工作表中所有的单元格。

## 示例

```javascript
/*本示例将 Sheet1 中单元格 C5 的字号设置为 14 磅。*/
function test() {
    Application.Worksheets.Item("Sheet1").Cells.Item(5, 3).Font.Size = 14
}
```

```javascript
/*本示例清除 Sheet1 上第一个单元格的公式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Cells.Item(1).ClearContents()
}
```

```javascript
/*本示例将 Sheet1 上所有单元格的字体设置为 8 磅“Arial”字体*/
function test() {
    Application.Worksheets.Item("Sheet1").Cells.Font.Name = "Arial"
    Application.Worksheets.Item("Sheet1").Cells.Font.Size = 8
}
```
