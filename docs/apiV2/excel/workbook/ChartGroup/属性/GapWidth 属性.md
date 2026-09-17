---
title: "GapWidth 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/GapWidth%20%E5%B1%9E%E6%80%A7.html
---
# GapWidth (属性)

对于条形图和柱形图：以条形或柱形宽度百分数的形式返回或设置条形簇或柱形簇之间的距离。对于饼图中的扇形和饼图中的条形：返回或设置图表的第一节和第二节之间的间距。**Long** 类型，可读写。

## 说明

该属性的取值必须介于 0 到 500 之间。

## 示例

```javascript
/*本示例将图表工作表 Chart1 第一个图表组的列簇间距设置为列宽的 50%。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.GapWidth = 50
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组中的列簇间距是否为 250 。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.GapWidth == 250)
}
```
