---
title: "Border 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
---
# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的模拟运算表的边框线条样式设置为虚线。*/
function test() {
    let datatable = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    datatable.Border.LineStyle = xlDash
}
```

```javascript
/*本示例为图表工作表 Chart1 中图表添加模拟运算表，并将该运算表边框设置为红色的细线。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasDataTable = true
    chart.DataTable.Border.ColorIndex = 3
    chart.DataTable.Border.Weight = xlHairline
}
```
