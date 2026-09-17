---
title: "Borders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
---
# Borders (属性)

返回或设置一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，它代表基于单元格边框格式的搜索条件。

## 示例

```javascript
/*本示例设置搜索条件以识别单元格的边框，这些单元格具有连续、加粗样式的底部边缘。创建一个满足该条件的单元格，然后找到该单元格并通知用户。*/
function test() {
    // Set the search criteria for the border of the cell format.
    let border = Application.FindFormat.Borders.Item(xlEdgeBottom)
    border.LineStyle = xlContinuous
    border.Weight = xlThick
    // Create a continuous thick bottom-edge border for cell A5.
    Application.Range("A5").Select()
    Application.Selection.Borders.Item(xlEdgeBottom).LineStyle = xlContinuous
    Application.Selection.Borders.Item(xlEdgeBottom).Weight = xlThick
    Application.Range("A1").Select()
    console.log("Cell A5 has a continuous thick bottom-edge border")
    // Find the cells based on the search criteria.
    Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false, null, true).Activate()
    console.log("ET has found this cell matching the search criteria.")
}
```

```javascript
/*本示例将替换条件中单元格底部边框的颜色设置为红色。*/
function test() {
    Application.ReplaceFormat.Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```
