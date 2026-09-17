---
title: "Depth 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/Depth%20%E5%B1%9E%E6%80%A7.html
---
# Depth (属性)

返回一个 **Long** 值，它代表分类轴刻度线标签的级别数。

## 示例

```javascript
/*下例显示 Sheet1 上的第一个图表分类轴刻度线标签的级别数。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.Depth)
}
```

```javascript
/*以下示例显示 Chart1 上分类轴刻度线标签的级别数。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(tickLabels.Depth)
}
```
