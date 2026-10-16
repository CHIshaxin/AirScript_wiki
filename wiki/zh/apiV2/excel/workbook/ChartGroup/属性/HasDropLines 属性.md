---
title: "HasDropLines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/HasDropLines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CCe7Y0mD
---

# HasDropLines (属性)

如果折线图或者面积图中有垂直线，则该值为 **True**。仅应用于折线图和面积图。**Boolean** 类型，可读写。

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
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组中是否有垂直线。本示例应在包含一个数据系列的二维折线图上运行。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.HasDropLines)
}
```
