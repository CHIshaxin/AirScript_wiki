---
title: "DownBars 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/DownBars%20%E5%B1%9E%E6%80%A7.html
---
# DownBars (属性)

返回一个 **[DownBars](./../../DownBars/DownBars 对象.html)** 对象，该对象表示折线图中的跌柱线。仅应用于折线图。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的涨跌柱线，并对其颜色进行设置。本示例应在包含两组彼此间有一个或多个相交数据点的数据系列的二维折线图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.DownBars.Interior.ColorIndex = 3
    chartgroup.UpBars.Interior.ColorIndex = 5
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表的第一个图表组的涨跌柱线，并将其边框颜色设置为蓝色。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.DownBars.Border.ColorIndex = 5
}
```
