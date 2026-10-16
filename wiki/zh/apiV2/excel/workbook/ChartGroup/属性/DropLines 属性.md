---
title: "DropLines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/DropLines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: m0JAXHqm
---

# DropLines (属性)

返回一个 **[DropLines](./../../DropLines/DropLines 对象.html)** 对象，该对象表示折线图或面积图上数据系列的垂直线。仅适用于折线图或面积图。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的垂直线，并对其线型、粗细和颜色进行设置。本示例应在包含一个数据系列的二维折线图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
    let border = chartgroup.DropLines.Border
    border.LineStyle = xlDash
    border.Weight = xlMedium
    border.ColorIndex = 3
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组的垂直线，并将其颜色设置为绿色。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
    chartgroup.DropLines.Border.Color = RGB(0, 255, 0)
}
```
