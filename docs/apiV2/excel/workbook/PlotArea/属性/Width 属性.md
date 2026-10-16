---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*本示例显示 Chart1 中绘图区的宽度。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Width)
}
```

```javascript
/*下例将 Sheet1 上第一个图表绘图区的宽度设置为400磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Width = 400
}
```
