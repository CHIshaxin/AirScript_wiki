---
title: "CellFormat (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/CellFormat%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B5I9gpzV
---

# CellFormat (对象)

代表单元格格式的搜索条件。

## 说明

使用 **[Application](./../Application/Application 对象.html)** 对象的 **FindFormat** 或 **ReplaceFormat** 属性可返回 **[CellFormat](./../CellFormat/CellFormat 对象.html)** 对象。

使用 **[CellFormat](./../CellFormat/CellFormat 对象.html)** 对象的 **Borders** 、**Font** 属性或 **[CellFormat](./../CellFormat/CellFormat 对象.html)** 对象的 **Interior** 属性可定义单元格格式的搜索条件。

## 示例

```javascript
/*本示例设置单元格格式内部的搜索条件。*/
function test() {
    // Set the interior of cell A1 to yellow.
    Application.Range("A1").Select()
    Application.Selection.Interior.ColorIndex = 36
    console.log("The cell format for cell A1 is a yellow interior.")

    // Set the CellFormat object to replace yellow with green.
    Application.FindFormat.Interior.ColorIndex = 36
    Application.ReplaceFormat.Interior.ColorIndex = 35

    // Find and replace cell A1's yellow interior with green.
    ActiveCell.Replace("", "", xlPart, xlByRows, false, null, true, true)
    console.log("The cell format for cell A1 is replaced with a green interior.")
}
```

```javascript
/*本示例将替换条件中单元格边框的粗细设置为 xlThick 。*/
function test() {
    Application.ReplaceFormat.Borders.Weight = xlThick
}
```
