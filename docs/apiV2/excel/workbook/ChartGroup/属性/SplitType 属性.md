---
title: "SplitType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/SplitType%20%E5%B1%9E%E6%80%A7.html
---
# SplitType (属性)

返回或设置在复合饼图或复合条饼图中拆分两部分的方式。**XlChartSplitType** 类型，可读写。

## 说明

| **XlChartSplitType** 可为以下 **XlChartSplitType** 常量之一。 |
| --- |
| **xlSplitByCustomSplit** |
| **xlSplitByPercentValue** |
| **xlSplitByPosition** |
| **xlSplitByValue** |

## 示例

```javascript
/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SplitType = xlSplitByValue
    chartgroup.SplitValue = 10
    chartgroup.VaryByCategories = true
}
```

```javascript
/*本示例以xlSplitByPosition拆分图表工作表 Chart1 的两部分，在第二部分中显示数据系列中的最小值。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SplitType = xlSplitByPosition
}
```
