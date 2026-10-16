---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DZoAjDk4
---

# Height (属性)

返回或设置一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例显示 Chart1 中的绘图区的高度。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Height)
}
```

```javascript
/*下例将 Sheet1 上的第一个图表绘图区的高度设置为300磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Height = 300
}
```
