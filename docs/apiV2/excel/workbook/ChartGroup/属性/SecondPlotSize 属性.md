---
title: "SecondPlotSize 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/SecondPlotSize%20%E5%B1%9E%E6%80%A7.html
---
# SecondPlotSize (属性)

以主饼图大小的百分比形式返回或设置复合饼图或复合条饼图中第二部分的大小。可为 5 到 200 之间的值。**Long** 类型，可读写。

## 示例

```javascript
/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。第二部分的大小为主条形图大小的百分之五十。*/
function test() {
    let chartobjects = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)
    chartobjects.SplitType = xlSplitByValue
    chartobjects.SplitValue = 10
    chartobjects.VaryByCategories = true
    chartobjects.SecondPlotSize = 50
}
```

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组中第二部分在主饼图的占比。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SecondPlotSize)
}
```
