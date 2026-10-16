---
title: "ShowNegativeBubbles 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/ShowNegativeBubbles%20%E5%B1%9E%E6%80%A7.html
---
# ShowNegativeBubbles (属性)

如果在图表组中显示表示负值的气泡，则该值为 **True**。仅对气泡图有效。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例设置工作表 Sheet1 第一个内嵌图表的第一个图表组表示负值的气泡为可见。仅对气泡图有效*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.ShowNegativeBubbles = true
}
```

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组表示负值的气泡是否可见。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.ShowNegativeBubbles)
}
```
