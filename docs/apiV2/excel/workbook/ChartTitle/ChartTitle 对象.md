---
title: "ChartTitle 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/ChartTitle%20%E5%AF%B9%E8%B1%A1.html
---
# ChartTitle (对象)

代表图表标题。

## 说明

使用 **ChartTitle** 属性可返回 **[ChartTitle](./../ChartTitle/ChartTitle 对象.html)** 对象。

只有图表的 **HasTitle** 属性为 **True** 时， **[ChartTitle](./../ChartTitle/ChartTitle 对象.html)** 对象才存在，从而才能使用该对象。

## 示例

```javascript
/*此示例为工作表 Sheet1 上第一个图表添加标题。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "February Sales"
}
```

```javascript
/*本示例将图表工作表 Chart1 的图表标题的字号设置为 20 磅。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.Font.Size = 20
}
```
