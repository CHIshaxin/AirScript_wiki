---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: _8fkMPgd
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 上第一个图表组的系列线粗细设置为3磅。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Format.Line.Weight = 3
}
```

```javascript
/*本示例将工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线的发光效果的半径设置为5。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Format.Glow.Radius = 5
}
```
