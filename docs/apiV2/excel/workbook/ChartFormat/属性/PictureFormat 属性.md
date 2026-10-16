---
title: "PictureFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/PictureFormat%20%E5%B1%9E%E6%80%A7.html
---
# PictureFormat (属性)

返回包含图片的指定图表的 **[PictureFormat](./../../PictureFormat/PictureFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例设置工作表 Sheet1 第一个内嵌图表的图表区图片亮度为 0.2 。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.PictureFormat.Brightness = 0.2
}
```

```javascript
/*此示例将图表工作表 Chart1 图表区的图片转换为黑白。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.PictureFormat.ColorType = msoPictureBlackAndWhite
}
```
