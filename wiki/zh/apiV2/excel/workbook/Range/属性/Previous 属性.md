---
title: "Previous (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Previous%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DQSfqweB
---

# Previous (属性)

返回一个代表上一个单元格的 **[Range](./../../Range/Range 对象.html)** 对象。

## 说明

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

## 示例

```javascript
/*此示例选定 sheet1 中上一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格左边的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Previous.Select()
}
```

```javascript
/*本示例将 B1 的上一个单元格的值设置为 a1，并将文本设置为斜体*/
function test() {
    let b1Previous = Range("B1").Previous
    b1Previous.Select()
    b1Previous.Font.Italic = true
}
```
