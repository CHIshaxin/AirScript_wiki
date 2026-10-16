---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LeaderLines/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DVbS0MJw
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 上的第二个数据系列的引导线粗细。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.LeaderLines.Border.Weight)
}
```

```javascript
/*本示例向第一张工作表上第一个图表的数据系列 3 添加线条样式为实线的引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。*/
function test() {
    let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3)
    Ser.HasLeaderLines = true
    Ser.LeaderLines.Border.LineStyle = xlContinuous
}
```
