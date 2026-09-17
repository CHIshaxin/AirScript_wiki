---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的第一个数据系列的误差线的线条样式设置为点划线。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).ErrorBars
    errorbars.Format.Line.DashStyle = msoLineDashDot
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表的第二个系列的误差线的阴影设置为不可见。*/
function test() {
    let errorbars = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    errorbars.Format.Shadow.Visible = false
}
```
