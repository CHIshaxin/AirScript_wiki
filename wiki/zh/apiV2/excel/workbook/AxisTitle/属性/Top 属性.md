---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 0cGi0eyb
---

# Top (属性)

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 中图表的分类轴标题到图表区顶部的距离。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    console.log(`分类轴到图表区顶部的距离：${axistitle.Top} 磅`)
}
```

```javascript
/*此示例将工作表 Sheet1 中第一张图表的数值轴标题到图表区顶部的距离设置为图表区高度的五分之一。*/
function test() {
    let chart = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart
    let axistitle = chart.Axes(xlValue).AxisTitle
    axistitle.Top = chart.ChartArea.Height / 5
}
```
