---
title: "DoughnutHoleSize (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/DoughnutHoleSize%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: hmiRmhCC
---

# DoughnutHoleSize (属性)

返回或设置圆环图图表组的内径大小。内径大小以图表大小的百分比表示，有效取值范围为 10% 到 90%。**Long** 类型，可读写。

## 示例

```javascript
/*本示例对图表工作表 Chart1 中第一个圆环图组的内径大小进行设置。本示例应在二维圆环图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.DoughnutHoleSize = 10
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个内嵌图表中第一个图表组的内径大小。本示例应在二维圆环图上运行。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.DoughnutHoleSize)
}
```
