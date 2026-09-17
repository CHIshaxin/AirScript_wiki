---
title: "ClearFormats 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
---
# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*此示例清除工作表 Sheet1 中第一张图表的第一个系列的误差线的格式设置。*/
function test() {
    let errorbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).ErrorBars
    errorbars.ClearFormats()
}
```

```javascript
/*此示例判断如果图表工作表 Chart1 中图表的第三个系列的误差线为红色，则清除该误差线的格式设置。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).ErrorBars
    if (errorbars.Border.ColorIndex == 3) {
        errorbars.ClearFormats()
    }
}
```
