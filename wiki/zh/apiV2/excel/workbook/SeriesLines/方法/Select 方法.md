---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DNeDrVeu
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*本示例选中图表工作表 Chart1 上第一个图表组的系列线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Select()
}
```

```javascript
/*本示例选中工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Select()
}
```
