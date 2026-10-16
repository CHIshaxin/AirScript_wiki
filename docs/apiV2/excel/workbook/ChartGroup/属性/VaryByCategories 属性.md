---
title: "VaryByCategories 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/VaryByCategories%20%E5%B1%9E%E6%80%A7.html
---
# VaryByCategories (属性)

如果 ET 对每个数据标志指定不同的颜色或模式，则该值为 **True**。图表中必须只包含一个数据系列。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例对图表工作表 Chart1 第一个图表组中的每个数据标志指定一种不同的颜色或图案。本示例应在数据系列上有数据标志的二维折线图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.VaryByCategories = true
}
```

```javascript
/*本示例显示工作表 Sheet1 第一个内嵌图表中第一个图表组中的每个数据标志是否指定一种不同的颜色或图案。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.VaryByCategories)
}
```
