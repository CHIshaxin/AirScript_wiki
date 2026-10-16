---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/HiLoLines/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: XlQRp1F1
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的高低点连线，并对其线型和颜色进行设置。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartGroup.HasHiLoLines = true
    let border = chartGroup.HiLoLines.Border
    border.LineStyle = xlDash
    border.ColorIndex = 4
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一个图表中第一个图表组的高低点连线的边框粗细是否为xlMedium。*/
function test() {
    let chartGroup = Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartGroup.HiLoLines.Border.Weight == xlMedium)
}
```
