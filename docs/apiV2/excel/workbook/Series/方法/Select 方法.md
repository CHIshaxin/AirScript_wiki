---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*此示例选择第一张工作表上的第一个图表的第二个数据系列。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series2.Select()
}
```

```javascript
/*此示例选择图表工作表 Chart1 上的第三个数据系列。*/
function test() {
    let series3 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    series3.Select()
}
```
