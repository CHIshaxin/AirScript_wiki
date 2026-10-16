---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CkmhxdYx
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的分类轴标题的边框设置为虚线。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Format.Line.DashStyle = msoLineDashDotDot
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴标题的前景色设置为蓝色。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    axistitle.Format.Fill.ForeColor.RGB = RGB(0, 255, 255)
}
```
