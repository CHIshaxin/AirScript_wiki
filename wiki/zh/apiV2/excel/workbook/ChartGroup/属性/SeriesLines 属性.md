---
title: "SeriesLines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/SeriesLines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BZVXWRqD
---

# SeriesLines (属性)

返回一个 SeriesLines 对象，该对象代表 2D 堆积条形图、2D 堆积柱形图、饼图的饼图或条形图的系列线。 此为只读属性。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表组一的系列线，并对其线型、粗细和颜色进行设置。本示例应在有两个或多个数据系列的二维堆积柱形图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    let border = chartgroup.SeriesLines.Border
    border.LineStyle = xlDash
    border.Weight = xlMedium
    border.ColorIndex = 3
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中的第一个图表组的系列线，并将其设置为蓝色。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    chartgroup.SeriesLines.Border.Color = RGB(0, 0, 255)
}
```
