---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: mCSNmPMq
---

# Orientation (属性)

返回或设置一个 **Long** 值，它代表文本方向。

## 说明

此属性值可设为 -90 到 90 度之间的整数值或 **XlTickLabelOrientation** 常量之一。

## 示例

```javascript
/*以下示例将 Chart1 上数值轴的刻度线标签的文本方向设置为向下排列。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    tickLabels.Orientation = xlTickLabelOrientationDownward
}
```

```javascript
/*下例显示 Sheet1 上的第一个图表分类轴刻度线标签的文本方向是否由 ET 设置。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.Orientation == xlTickLabelOrientationAutomatic)
}
```
