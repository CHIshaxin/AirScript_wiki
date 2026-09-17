---
title: "Overlap 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/Overlap%20%E5%B1%9E%E6%80%A7.html
---
# Overlap (属性)

指定放置条形和柱形的方式。取值范围在 -100 到 100 之间。仅应用于二维条形图和二维柱形图。**Long** 类型，可读写。

## 说明

如果将该属性设置为 -100，条形之间有一个条形的间距。如果重叠比例为 0（零），条形之间没有间距（一个条形紧挨着另一条形）。如果重叠比例为 100，所有条形将完全重叠。

## 示例

```javascript
/*本示例将图表工作表 Chart1 第一个图表组的重叠比例设置为 -50 。本示例应在包含两个或更多系列的二维柱形图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.Overlap = -50
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组的重叠比例。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.Overlap)
}
```
