---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LeaderLines/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CqlN0o4d
---

# Select (方法)

选择对象。

## 示例

```javascript
/*此示例选中图表工作表 Chart1 上的第一个数据系列的引导线。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series1.LeaderLines.Select()
}
```

```javascript
/*本示例选中第一张工作表上第一个图表的数据系列 1 的引导线。*/
function test() {
    let series1 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series1.LeaderLines.Select()
}
```
