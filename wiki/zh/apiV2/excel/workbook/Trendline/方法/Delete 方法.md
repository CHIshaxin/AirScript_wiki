---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DIumygTv
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除 Chart1 中第一个数据系列索引为1的趋势线。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1)
    series1.Trendlines(1).Delete()
}
```

```javascript
/*此示例删除第一张工作表上的第一个图表的第二个数据系列索引为1的趋势线。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series2.Trendlines(1).Delete()
}
```
