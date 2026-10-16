---
title: "Left (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: SAB0MEz1
---

# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的图表区距离左边缘 100 磅。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Left = 100
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表的图表区距离左边缘的距离。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Left)
}
```
