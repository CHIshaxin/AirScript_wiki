---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D71rgZTt
---

# Height (属性)

返回或设置一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的高度为 650 磅。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Height = 650
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个内嵌图表的高度。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.Height)
}
```
