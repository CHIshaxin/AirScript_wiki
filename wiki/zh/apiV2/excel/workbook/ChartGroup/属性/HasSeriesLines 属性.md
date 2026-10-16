---
title: "HasSeriesLines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/HasSeriesLines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bh_Ec_Ej
---

# HasSeriesLines (属性)

如果堆积柱形图或堆积条形图有系列线，或者复合饼图或复合条饼图在两个部分之间有连接线，则该属性值为 **True**。仅适用于二维堆积柱形图、堆积条形图、复合饼图或复合条饼图。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例打开图表工作表 Chart1 上第一个图表组的系列线，然后设置其线条样式、粗细和颜色。该示例应在具有两个或更多系列的 2D 堆积柱形图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    let border = chartgroup.SeriesLines.Border
    border.LineStyle = xlContinuous
    border.Weight = xlMedium
    border.ColorIndex = 3
}
```

```javascript
/*本示例打开工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线，并显示其是否为细线。该示例应在具有两个或更多系列的 2D 堆积柱形图上运行。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    console.log(chartgroup.SeriesLines.Border.Weight == xlHairline)
}
```
