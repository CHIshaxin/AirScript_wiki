---
title: "ErrorBars 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/ErrorBars%20%E5%AF%B9%E8%B1%A1.html
---
# ErrorBars (对象)

代表图表数据系列上的误差线。

## 说明

误差线表明图表数据的不确定程度。只有二维图表上的面积、条形、柱形、折线和散点图组中的数据系列可以有误差线。只有散点图组中的数据系列可以有 x 误差线和 y 误差线。此对象不是集合。没有代表单个误差线的对象；或者打开系列中所有数据点的 x 误差线或 y 误差线，或者将其全部关闭。

**ErrorBar**方法更改误差线的格式和类型。

使用 **ErrorBars** 属性可返回 **[ErrorBars](./../ErrorBars/ErrorBars 对象.html)** 对象。

## 示例

```javascript
/*本示例打开工作表 Sheet1 嵌入的第一个图表的第一个数据系列的误差线，并设置误差线的尾部样式。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).HasErrorBars = true
    ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个系列的误差线设置为黄色虚线。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    errorbars.Border.ColorIndex = 6
    errorbars.Border.LineStyle = xlDash
}
```
