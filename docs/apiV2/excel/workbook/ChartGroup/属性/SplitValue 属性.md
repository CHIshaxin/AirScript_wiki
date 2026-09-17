---
title: "SplitValue 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/SplitValue%20%E5%B1%9E%E6%80%A7.html
---
# SplitValue (属性)

返回或设置复合饼图或复合条饼图中分隔两部分的阈值。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例必须在复合饼图或复合条饼图中运行。本示例在工作表 Sheet1 第一个内嵌图表的第一个图表组中，以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SplitType = xlSplitByValue
    chartgroup.SplitValue = 10
    chartgroup.VaryByCategories = true
}
```

```javascript
/*本示例显示图表工作表 Chart1 中分隔饼图的两个部分的阈值。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SplitValue)
}
```
