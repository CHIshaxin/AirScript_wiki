---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/HiLoLines/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的高低点连线的前景色是否为红色。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    let chartFormat = chartGroup.HiLoLines.Format
    console.log(chartFormat.Fill.ForeColor.RGB == RGB(255, 0, 0))
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一个图表中第一个图表组的高低点连线的虚线样式。*/
function test() {
    let chartGroup = Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartGroup.HiLoLines.Format.Line.DashStyle = msoLineRoundDot
}
```
