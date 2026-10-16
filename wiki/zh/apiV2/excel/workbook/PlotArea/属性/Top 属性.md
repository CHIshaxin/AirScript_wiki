---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: eH48aSGf
---

# Top (属性)

返回或设置一个 **Double**值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*本示例显示 Chart1 中的绘图区上边缘到图表上的图表区顶部的距离。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Top)
}
```

```javascript
/*下例将 Sheet1 上第一个图表的绘图区上边缘到图表上的图表区顶部的距离设置为50磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Top = 50
}
```
