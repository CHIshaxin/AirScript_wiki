---
title: "ChartType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/ChartType%20%E5%B1%9E%E6%80%A7.html
---
# ChartType (属性)

返回或设置图表类型。 **[XlChartType](./../../枚举/XlChartType 枚举.html)** 类型，可读写。

## 说明

某些图表类型对数据透视图报表无效。

## 示例

```javascript
/*此示例显示活动工作表上的第一个图表的第一个系列的图表类型是否为气泡图。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).ChartType == xlBubble)
}
```

```javascript
/*此示例将图表工作表 Chart1 上的第一个数据系列的图表类型设置为折线散点图。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series1.ChartType = xlXYScatterLines
}
```
