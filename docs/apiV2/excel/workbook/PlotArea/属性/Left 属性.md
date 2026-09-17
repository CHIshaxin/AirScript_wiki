---
title: "Left 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
---
# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例显示 Chart1 中的绘图区左边缘到图表上的图表区左边缘的距离。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Left)
}
```

```javascript
/*下例将 Sheet1 上第一个图表的绘图区左边缘到图表上的图表区左边缘的距离设置为10磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Left = 10
}
```
