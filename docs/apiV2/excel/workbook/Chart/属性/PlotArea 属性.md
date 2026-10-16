---
title: "PlotArea 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/PlotArea%20%E5%B1%9E%E6%80%A7.html
---
# PlotArea (属性)

返回一个 **[PlotArea](./../../PlotArea/PlotArea 对象.html)** 对象，该对象表示图表的绘图区。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 绘图区的内部颜色设置为蓝绿色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.PlotArea.Interior.ColorIndex = 8
}
```

```javascript
/*本示例将工作表Sheet1第一张内嵌图表的图表区设置为虚线边框，绘图区设置为点线边框。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartArea.Border.LineStyle = xlDash
    chart.PlotArea.Border.LineStyle = xlDot
}
```
