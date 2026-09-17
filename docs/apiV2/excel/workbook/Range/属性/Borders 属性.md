---
title: "Borders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
---
# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

## 示例

```javascript
/*此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框。*/
function test() {
    let b2_bottom = Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)
    b2_bottom.LineStyle = xlContinuous
    b2_bottom.Weight = xlThin
    b2_bottom.ColorIndex = 3
}
```

```javascript
/*本示例将 A1:E3 单元格区域的边框设置为绿色、中等粗细，并使颜色变暗*/
function test() {
    Range("A1:E3").Borders.Color = RGB(0, 255, 0)
    Range("A1:E3").Borders.Weight = xlMedium
    Range("A1:E3").Borders.TintAndShade = -1
}
```
