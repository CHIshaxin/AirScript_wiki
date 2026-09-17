---
title: "ChartArea 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/ChartArea%20%E5%AF%B9%E8%B1%A1.html
---
# ChartArea (对象)

代表图表的图表区。

## 说明

图表区包含绘图区在内的一切内容。但是，绘图区具有它自己的填充方式，因此填充绘图区并不会填充图表区。

有关设置绘图区格式的信息，请参阅 **PlotArea 对象**。

使用 **ChartArea** 属性可返回 **[ChartArea](./../ChartArea/ChartArea 对象.html)** 对象。

## 示例

```javascript
/*本示例关闭 Sheet1 工作表上第一个嵌入图表的图表区边框。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Line.Visible = false
}
```

```javascript
/*本示例将工作表 Chart1 图表区的内部颜色设置为蓝色。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Fill.ForeColor.RGB = rgbBlue
}
```
