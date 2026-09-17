---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回对象的宽度（以磅为单位）。只读。

## 返回值

Double

## 示例

```javascript
/*此示例显示第一张工作表上第一个图表的第二个数据系列中第二个数据点的宽度。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Points(2).Width)
}
```

```javascript
/*本示例显示 Chart1 中第三个数据系列的第三个数据点的宽度。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series.Points(3).Width)
}
```
