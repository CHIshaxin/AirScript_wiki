---
title: "PlotArea (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/PlotArea%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BcNBwFFB
---

# PlotArea (对象)

代表图表的绘图区。

## 说明

该区域为绘制图表数据的区域。二维图表中的绘图区包含数据标志、网格线、数据标签、趋势线和可选的置于图表区内的图表项。三维图表的绘图区中除包含上述各项外，还在图表中包含背景墙、基底、坐标轴、坐标轴标题和刻度线标签。绘图区被图表区所包围。二维图表的图表区包含坐标轴、图表标题、坐标轴标题和图例。三维图表的图表区包含图表标题和图例。有关设置图表区格式的详细信息，请参阅 ChartArea 对象。

## 示例

```javascript
/*本示例将 Chart1 中的绘图区前景色设为蓝色。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    plotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
}
```

```javascript
/*下例将 Sheet1 上的第一个图表绘图区内部宽度设置为400磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.InsideWidth = 400
}
```
