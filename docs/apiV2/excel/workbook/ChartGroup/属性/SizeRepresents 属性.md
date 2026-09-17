---
title: "SizeRepresents 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/SizeRepresents%20%E5%B1%9E%E6%80%A7.html
---
# SizeRepresents (属性)

返回或设置气泡图中气泡的大小所表示的含义。可为以下 **XlSizeRepresents** 常量之一：**xlSizeIsArea** 或 **xlSizeIsWidth**。**Long** 类型，可读写。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 第一个图表组中气泡的大小所表示的含义。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SizeRepresents = xlSizeIsWidth
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组中气泡的大小是否表示为气泡的面积。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SizeRepresents == xlSizeIsArea)
}
```
