---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LeaderLines/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DrWWXi0j
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例将图表工作表 Chart1 上的第一个数据系列的引导线发光半径设置为5磅。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series1.LeaderLines.Format.Glow.Radius = 5
}
```

```javascript
/*本示例将第一张工作表上第一个图表的数据系列 1 的引导线的粗细设置为2磅。*/
function test() {
    let series1 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series1.LeaderLines.Format.Line.Weight = 2
}
```
