---
title: "Forward2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Forward2%20%E5%B1%9E%E6%80%A7.html
---
# Forward2 (属性)

返回或设置趋势线向前延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。

## 示例

```javascript
/*本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含具有趋势线的单个序列的 2D 柱形图上运行*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    trendline.Forward2 = 5
    trendline.Backward2 = 0.5
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第一个数据系列索引为2的趋势线向前延伸的单位数。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series.Trendlines(2).Forward2)
}
```
