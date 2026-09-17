---
title: "Border 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DropLines/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
---
# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例打开工作表 Sheet1 中第一个图表的第一个图表组的垂直线，并将该垂直线设置为虚线。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
    chartgroup.DropLines.Border.LineStyle = xlDash
}
```

```javascript
/*本示例打开图表工作表 Chart1 中图表的第一个图表组的垂直线，并将该垂直线的颜色设置为黄色。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
    chartgroup.DropLines.Border.ColorIndex = 6
}
```
