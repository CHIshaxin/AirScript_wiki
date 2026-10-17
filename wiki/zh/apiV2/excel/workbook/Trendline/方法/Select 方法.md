---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: RJqcVMJE
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*下例选中 Chart1 中的第二个数据系列索引为1的趋势线。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Trendlines(1)
    trendline.Select()
}
```

```javascript
/*此示例选中第一张工作表上的第一个图表的第一个数据系列索引为2的趋势线。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series.Trendlines(2).Select()
}
```
