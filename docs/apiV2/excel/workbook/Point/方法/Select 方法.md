---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*此示例选中第一张工作表上第一个图表的第一个数据系列中第7个数据点。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series.Points(7).Select()
}
```

```javascript
/*本示例选中 Chart1 中第三个数据系列的第8个数据点。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    series.Points(8).Select()
}
```
