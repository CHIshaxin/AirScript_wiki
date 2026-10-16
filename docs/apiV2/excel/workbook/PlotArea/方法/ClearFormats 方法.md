---
title: "ClearFormats 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
---
# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*本示例清除 Chart1 中绘图区的格式设置。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    plotArea.ClearFormats()
}
```

```javascript
/*下例清除 Sheet1 上第一个图表绘图区的格式设置。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.ClearFormats()
}
```
