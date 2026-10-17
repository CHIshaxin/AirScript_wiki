---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DSeHE7Kv
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*下例更改 Chart1 中的第二个数据系列索引为1的趋势线的虚线样式。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Trendlines(1)
    trendline.Format.Line.DashStyle = msoLineLongDash
}
```

```javascript
/*此示例设置第一张工作表上的第一个图表的第一个数据系列索引为2的趋势线发光半径为2。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series.Trendlines(2).Format.Glow.Radius = 2
}
```
