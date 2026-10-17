---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BLLXrDv_
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*下例将 Sheet1 上的第一个图表分类轴刻度线标签的前景色设为蓝色。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    tickLabels.Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
}
```

```javascript
/*以下示例将 Chart1 上数值轴刻度线标签的发光半径设置为2磅。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    tickLabels.Format.Glow.Radius = 2
}
```
