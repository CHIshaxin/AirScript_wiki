---
title: "PictureType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/PictureType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BDQb7Std
---

# PictureType (属性)

返回或设置一个 **XlChartPictureType**值，它代表在柱形或条形图上图片的显示方式。

## 示例

```javascript
/*本示例设置 Chart1 中的系列一伸展图片。本示例应在带图片数据标志的二维柱形图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).PictureType = xlStretch
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列的图片是否为堆叠类型。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.PictureType == xlStack)
}
```
