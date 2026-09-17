---
title: "EndStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E5%B1%9E%E6%80%A7/EndStyle%20%E5%B1%9E%E6%80%A7.html
---
# EndStyle (属性)

返回或设置误差线的尾端样式。可为以下 **XlEndStyleCap** 常量之一：**xlCap** 或 **xlNoCap**。**Long** 类型，可读写。

## 示例

```javascript
/*本示例对 Chart1 中第一个系列的误差线的尾端样式进行设置。本示例必须在其第一个系列带 Y 误差线的二维折线图上运行。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).ErrorBars
    errorbars.EndStyle = xlCap
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的第二个系列的误差线的尾端样式是否不应用帽。*/
function test() {
    let errorbars = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    console.log(errorbars.EndStyle == xlNoCap)
}
```
