---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Gridlines/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 数值轴的主要网格线的前景色是否为蓝色。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    let chartFormat = axis.MajorGridlines.Format
    console.log(chartFormat.Fill.ForeColor.RGB == RGB(0, 0, 255))
}
```

```javascript
/*本示例打开工作表 Sheet1 中第一个图表类别轴的次要网格线，并设置网格线的虚线样式。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasMinorGridlines = true
    axis.MinorGridlines.Format.Line.DashStyle = msoLineDash
}
```
