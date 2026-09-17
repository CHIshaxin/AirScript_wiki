---
title: "InsideHeight 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/InsideHeight%20%E5%B1%9E%E6%80%A7.html
---
# InsideHeight (属性)

以磅为单位返回绘图区内部高度。可读写 **Double** 类型。

## 说明

这种度量方式的图形区域不包含坐标轴标签。图形区域的 Height 属性使用包含坐标轴标签的封闭矩形。

## 示例

```javascript
/*本示例在 Chart1 中的绘图区内绘制带点线的矩形。*/
function test() {
    let chart = Application.Charts("Chart1").ChartObjects(1).Chart
    let plotArea = chart.PlotArea
    let shape = chart.Shapes.AddShape(msoShapeRectangle, plotArea.InsideLeft, plotArea.InsideTop, plotArea.InsideWidth, plotArea.InsideHeight)
    shape.Fill.Transparency = 1
    shape.Line.DashStyle = msoLineDashDot
}
```

```javascript
/*下例显示 Sheet1 上的第一个图表绘图区内部高度。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.InsideHeight)
}
```
