---
title: "Chart (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/Chart%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: KitmdalX
---

# Chart (属性)

返回一个 **[Chart](./../../Chart/Chart 对象.html)** 对象，该对象代表对象中包含的图表。只读。

## 示例

```javascript
/*本示例为工作表 Sheet1 上嵌入的第一个图表添加标题。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "1995 Rainfall Totals by Month"
}
```

```javascript
/*本示例显示图表工作表 Chart1 的图表样式。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    console.log(chart.ChartStyle)
}
```
