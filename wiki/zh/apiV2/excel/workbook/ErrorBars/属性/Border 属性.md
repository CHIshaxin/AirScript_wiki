---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: H7vxqug5
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*此示例将工作表 Sheet1 中第一张图表的第一个序列的误差线设置为虚线。*/
function test() {
    let errorbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).ErrorBars
    errorbars.Border.LineStyle = xlDash
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个序列的误差线设置为蓝色细线。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    errorbars.Border.ColorIndex = 5
    errorbars.Border.Weight = xlHairline
}
```
