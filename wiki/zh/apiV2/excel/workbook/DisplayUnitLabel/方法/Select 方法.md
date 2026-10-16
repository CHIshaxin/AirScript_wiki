---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BrHOI6KG
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*本示例选中图表工作表 Chart1 中数值轴上的单位标签。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Select()
}
```

```javascript
/*下例选中 Sheet1 上的第一个图表数值轴上的单位标签。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Select()
}
```
