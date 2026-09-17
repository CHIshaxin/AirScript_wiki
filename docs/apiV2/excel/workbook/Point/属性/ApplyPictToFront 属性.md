---
title: "ApplyPictToFront 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/ApplyPictToFront%20%E5%B1%9E%E6%80%A7.html
---
# ApplyPictToFront (属性)

如果图片应用于系列中数据点或所有数据点之前，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例将图片应用于图表工作表 Chart1 上第二个系列的第三个数据点之前。图片必须已应用于该系列（设置此属性将更改图片的方向）。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(3)
    point.ApplyPictToFront = true
}
```

```javascript
/*此示例判断图片是否应用于图表工作表 Chart1 上第一个系列的第三个数据点之前。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(3)
    if (point.ApplyPictToFront) {
        console.log("图片应用于数据点之前")
    } else {
        console.log("图片没有应用于数据点之前")
    }
}
```
