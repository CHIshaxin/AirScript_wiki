---
title: "PlotBy (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/PlotBy%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B6phK2vv
---

# PlotBy (属性)

返回或设置行或列在图表中作为数据系列使用的方式。可为以下 **XlRowCol** 常量之一：**xlColumns** 或 **xlRows**。**Long** 类型，可读写。

## 说明

对于数据透视图，该属性是只读的，并且总返回 **xlColumns**。

## 示例

```javascript
/*本示例使第一张工作表第一张嵌入图表按列绘制数据。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.PlotBy = xlColumns
}
```

```javascript
/*本示例显示图表工作表 Chart1 是否按列绘制数据。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.PlotBy == xlColumns)
}
```
