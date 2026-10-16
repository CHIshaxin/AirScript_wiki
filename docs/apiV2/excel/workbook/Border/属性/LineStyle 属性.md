---
title: "LineStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Border/%E5%B1%9E%E6%80%A7/LineStyle%20%E5%B1%9E%E6%80%A7.html
---
# LineStyle (属性)

返回或设置边框的线型。 **[XlLineStyle](./../../枚举/XlLineStyle 枚举.html)** 、**xlGray25**、**xlGray50**、**xlGray75** 或 **xlAutomatic** 类型，可读写。

## 说明

**xlDouble** 和 **xlSlantDashDot** 不适用于图表。

## 示例

```javascript
/*本示例为 Chart1 的图表区和绘图区域设置边框*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.ChartArea.Border.LineStyle = xlDashDot
    let border = chart.PlotArea.Border
    border.LineStyle = xlDashDotDot
    border.Weight = xlThick
}
```

```javascript
/*本示例将 H6 单元格的顶部的边框线样式设置为虚线。*/
function test() {
    Range("H6").Borders.Item(xlEdgeTop).LineStyle = xlDash
}
```
