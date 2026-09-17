---
title: "PictureType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/PictureType%20%E5%B1%9E%E6%80%A7.html
---
# PictureType (属性)

返回或设置一个 **XlChartPictureType** 值，它代表在柱形或条形图上图片的显示方式。

## 示例

```javascript
/*此示例设置第一张工作表上第一个图表的第三个数据系列中第2个数据点图片的显示方式为xlStack。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3)
    series.Points(2).PictureType = xlStack
}
```

```javascript
/*本示例显示 Chart1 中第二个数据系列的第9个数据点图片的显示方式是否为xlStretch。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Points(9).PictureType == xlStretch)
}
```
