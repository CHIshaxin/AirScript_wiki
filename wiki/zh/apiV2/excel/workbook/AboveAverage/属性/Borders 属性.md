---
title: "Borders (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CK4-zR9A
---

# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。

## 说明

对于条件格式对象，您只能为单元格的上边框、下边框和侧边框设置属性。

## 示例

```javascript
/*此示例将工作表 Sheet1 中区域 B1:B10 的第一个（AboveAverage）条件格式的底部边框颜色设置为红色细边框。*/
function test() {
    let aboveAverage = Application.Worksheets.Item("Sheet1").Range("B1:B10").FormatConditions.Item(1)
    let boder = aboveAverage.Borders.Item(xlEdgeBottom)
    boder.LineStyle = xlContinuous
    boder.Weight = xlThin
    boder.ColorIndex = 3
}
```

```javascript
/*本示例设置活动工作表上区域 B1:B10 中第一个（AboveAverage）条件格式边框的线条样式和粗细。*/
function test() {
    let aboveAverage = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    let boders = aboveAverage.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThick
}
```
