---
title: "ChartType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ChartType%20%E5%B1%9E%E6%80%A7.html
---
# ChartType (属性)

返回或设置图表类型。 **[XlChartType](./../../枚举/XlChartType 枚举.html)** 类型，可读写。

## 说明

某些图表类型对数据透视图报表无效。

## 示例

```javascript
/*本示例判断如果第一张工作表第一张内嵌图表为二维气泡图表，则将第一个图表组中的气泡大小设置为默认大小的 200%。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    if (chart.ChartType == xlBubble) {
        chart.ChartGroups(1).BubbleScale = 200
    }
}
```

```javascript
/*本示例修改图表工作表 Chart1 为百分比堆积条形图。*/
function test() {
    let charts = Application.Charts.Item("Chart1")
    charts.ChartType = xlBarStacked100
}
```
