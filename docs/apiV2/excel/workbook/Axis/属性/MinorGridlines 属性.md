---
title: "MinorGridlines 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinorGridlines%20%E5%B1%9E%E6%80%A7.html
---
# MinorGridlines (属性)

返回 **[Gridlines](./../../Gridlines/Gridlines 对象.html)** 对象，该对象表示指定坐标轴的次要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

## 示例

```javascript
/*本示例对图表工作表 Chart1 数值轴的次要网格线的颜色进行设置。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    if (axis.HasMinorGridlines) {
        //set color to blue
        axis.MinorGridlines.Border.ColorIndex = 5
    }
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表的类别轴的次要网格线的设置为红色虚线。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.MinorGridlines.Border.ColorIndex = 3
    axis.MinorGridlines.Border.LineStyle = xlDash
}
```
