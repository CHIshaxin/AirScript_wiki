---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BClmlEQD
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示 Chart1 中的绘图区名称。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Name)
}
```

```javascript
/*下例显示 Sheet1 上第一个图表的绘图区名称。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    console.log(plotArea.Name)
}
```
