---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dm6zTmq4
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。只读。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表区的名称。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Name)
}
```

```javascript
/*本示例显示图表工作表 Chart1 的图表区名称是否为图表。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Name == "图表")
}
```
