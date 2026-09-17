---
title: "RadarAxisLabels 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/RadarAxisLabels%20%E5%B1%9E%E6%80%A7.html
---
# RadarAxisLabels (属性)

返回一个 **[TickLabels](./../../TickLabels/TickLabels 对象.html)** 对象，该对象表示指定图表组的雷达图坐标轴标签。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的雷达图坐标轴标签，并对这些标签的颜色进行设置。本示例应在雷达图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasRadarAxisLabels = true
    chartgroup.RadarAxisLabels.Font.ColorIndex = 3
}
```

```javascript
/*本示例将工作表 Sheet1 第一个内嵌图表中第一个图表组的雷达图坐标轴标签启用，并显示其字号。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasRadarAxisLabels = true
    console.log(chartgroup.RadarAxisLabels.Font.Size)
}
```
