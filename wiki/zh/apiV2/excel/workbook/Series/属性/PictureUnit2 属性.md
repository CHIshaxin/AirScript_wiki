---
title: "PictureUnit2 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/PictureUnit2%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: fwyYwpF3
---

# PictureUnit2 (属性)

如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

## 示例

```javascript
/*此示例将 Chart1 中的系列一设为堆积图片，设置每一图片代表五个单位。此示例应在带图片数据标志的二维柱形图上运行。*/
function test() {
    let Ser = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    Ser.PictureType = xlStackScale
    Ser.PictureUnit2 = 5
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第三个数据系列的图片单位。此示例应在带图片数据标志的二维柱形图上运行，且图片类型为xlStackScale。*/
function test() {
    let series3 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series3.PictureUnit2)
}
```
