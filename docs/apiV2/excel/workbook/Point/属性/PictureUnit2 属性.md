---
title: "PictureUnit2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/PictureUnit2%20%E5%B1%9E%E6%80%A7.html
---
# PictureUnit2 (属性)

如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

## 示例

```javascript
/*此示例设置第一张工作表上第一个图表的第二个数据系列的第4个数据点上每个图片的单位。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series.Points(4).PictureUnit2 = 5
}
```

```javascript
/*本示例显示 Chart1 中第一个数据系列的第4个数据点上每个图片的单位。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series.Points(4).PictureUnit2)
}
```
