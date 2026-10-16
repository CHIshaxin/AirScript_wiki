---
title: "Position 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
---
# Position (属性)

返回或设置图表上绘制区域的位置。可读/写 **XlChartElementPosition**类型。

## 示例

```javascript
/*本示例显示 Chart1 中的绘图区位置是否为xlChartElementPositionAutomatic。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Position == xlChartElementPositionAutomatic)
}
```

```javascript
/*下例将 Sheet1 上第一个图表的绘图区位置设置成xlChartElementPositionCustom。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Position = xlChartElementPositionCustom
}
```
