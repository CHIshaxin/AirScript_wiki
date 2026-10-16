---
title: "ChartColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ChartColor%20%E5%B1%9E%E6%80%A7.html
---
# ChartColor (属性)

返回或设置一个 Integer 类型的值，该值代表图表的配色方案。 读/写。

## 示例

```javascript
/*本示例设置工作表 Sheet1 中第一个内嵌图表的配色方案。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartColor = 1
}
```

```javascript
/*本示例显示图表工作表 Chart1 的配色方案。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    console.log(chart.ChartColor)
}
```
