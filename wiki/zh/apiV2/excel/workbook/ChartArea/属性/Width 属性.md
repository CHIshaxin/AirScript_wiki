---
title: "Width (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: kaVXlNsP
---

# Width (属性)

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表图表区的宽度。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Width)
}
```

```javascript
/*此示例设置图表工作表 Chart1 图表区的宽度为 600 磅。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Width = 600
}
```
