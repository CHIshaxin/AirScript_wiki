---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BaBnT5ry
---

# Top (属性)

返回一个值，该值代表从对象上边缘到图表区上边缘的距离（以磅为单位）。只读。

## 返回值

Double

## 示例

```javascript
/*此示例显示第一张工作表上第一个图表的第二个数据系列中第4个数据点的上边缘到图表区上边缘的距离。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Points(4).Top)
}
```

```javascript
/*本示例显示 Chart1 中第一个数据系列的第三个数据点的上边缘到图表区上边缘的距离。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series.Points(3).Top)
}
```
