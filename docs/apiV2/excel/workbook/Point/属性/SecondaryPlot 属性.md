---
title: "SecondaryPlot 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/SecondaryPlot%20%E5%B1%9E%E6%80%A7.html
---
# SecondaryPlot (属性)

如果数据点位于复合饼图或复合条饼图中的第二个饼图或条形图上，则该值为 **True**。仅适用于复合饼图或复合条饼图中的数据点。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例必须在饼图或饼图条形图上运行。 本示例将第四个数据点移到图表的第二部分。*/
function test() {
    Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Points(4).SecondaryPlot = true
}
```

```javascript
/*本示例显示 Chart1 中第一个数据系列的第3个数据点是否在复合条饼图中的第二个饼图上。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series.Points(3).SecondaryPlot)
}
```
