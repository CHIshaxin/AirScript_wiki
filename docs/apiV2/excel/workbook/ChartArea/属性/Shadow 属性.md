---
title: "Shadow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
---
# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*此示例将图表工作表 Chart1 的图表区设置为有阴影。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Shadow = true
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表的图表区是否有阴影。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Shadow)
}
```
