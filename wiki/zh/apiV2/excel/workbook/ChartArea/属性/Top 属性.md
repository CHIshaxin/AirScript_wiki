---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D1xcr2-Z
---

# Top (属性)

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表的图表区到工作表顶部的距离。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Top)
}
```

```javascript
/*此示例设置图表工作表 Chart1 图表区到顶部的距离为 50 磅。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Top = 50
}
```
