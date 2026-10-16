---
title: "MajorGridlines (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MajorGridlines%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: __Ti4K2c
---

# MajorGridlines (属性)

返回一个 **[Gridlines](./../../Gridlines/Gridlines 对象.html)** 对象，该对象表示指定坐标轴的主要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

## 示例

```javascript
/*本示例为图表工作表 Chart1 中数值轴的主要网格线设置为蓝色。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlValue)
    if (axis.HasMajorGridlines) {
        //set color to blue
        axis.MajorGridlines.Border.ColorIndex = 5
    }
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴的主要网格线设置为红色的虚线。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.MajorGridlines.Border.ColorIndex = 3
    axis.MajorGridlines.Format.Line.DashStyle = msoLineDashDotDot
}
```
