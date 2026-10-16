---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 返回值

Varint

## 示例

```javascript
/*本示例选中 Chart1 中的绘图区。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    plotArea.Select()
}
```

```javascript
/*下例选中 Sheet1 上第一个图表的绘图区。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Select()
}
```
