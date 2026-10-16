---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组的索引号。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.Index)
}
```

```javascript
/*本示例判断如果图表工作表 Chart1 中第一个图表组的索引号为1，则开启该图表组中的垂直线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    if (chartgroup.Index == 1) {
        chartgroup.HasDropLines = true
    }
}
```
