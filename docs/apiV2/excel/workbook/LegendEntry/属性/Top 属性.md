---
title: "Top 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
---
# Top (属性)

返回一个 **Double**值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第三个图例项上边缘到图表区顶部的距离。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(3).Top)
}
```

```javascript
/*下例显示 Chart1 上第二个图例项上边缘到图表区顶部的距离。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).Top)
}
```
