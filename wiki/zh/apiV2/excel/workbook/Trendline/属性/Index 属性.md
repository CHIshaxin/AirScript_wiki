---
title: "Index (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: PRSRtyBT
---

# Index (属性)

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*下例显示 Chart1 中的第三个数据系列指定趋势线的索引。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Trendlines(1)
    console.log(trendline.Index)
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第一个数据系列指定趋势线的索引。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series.Trendlines(2).Index)
}
```
