---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DRovz2rs
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*以下示例选中 Chart1 上数值轴上的刻度线标志。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    tickLabels.Select()
}
```

```javascript
/*下例选中 Sheet1 上的第一个图表分类轴上的刻度线标志。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    tickLabels.Select()
}
```
