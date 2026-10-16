---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 48drYh46
---

# Top (属性)

返回或设置一个 **Double**值，它代表从对象的上边缘到图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表标题设置为可见，并显示标题到图表顶部的距离。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasTitle = true
    console.log(`标题到顶部的距离为 ${chart.ChartTitle.Top} 磅`)
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的图表标题到图表顶部的距离是否为 6 磅。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    console.log(charttitle.Top == 6)
}
```
