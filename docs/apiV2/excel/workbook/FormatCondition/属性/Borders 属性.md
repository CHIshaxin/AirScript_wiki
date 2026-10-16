---
title: "Borders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
---
# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

## 示例

```javascript
/*此示例将工作表 Sheet1 中区域 A1:A10 的第一个条件格式的底部边框颜色设置为红色细边框。*/
function test() {
    let formatCondition = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    let boder = formatCondition.Borders.Item(xlEdgeBottom)
    boder.LineStyle = xlContinuous
    boder.Weight = xlThin
    boder.ColorIndex = 3
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式边框的线条样式和粗细。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    let boders = databar.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThick
}
```
