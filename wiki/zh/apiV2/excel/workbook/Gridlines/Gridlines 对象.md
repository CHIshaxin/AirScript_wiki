---
title: "Gridlines (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Gridlines/Gridlines%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: 5pN8GFB7
---

# Gridlines (对象)

代表图表坐标轴的主要和次要网格线。

## 说明

网格线延伸图表坐标轴上的刻度线，以便更容易地分辨与数据标志相关联的数值。此对象不是集合。同时也没有代表单个网格线的对象；要么打开坐标轴上所有的网格线，要么将其全部关闭。

使用 **MajorGridlines** 属性可返回 **GridLines** 对象，该对象代表坐标轴的主要网格线。使用 **MinorGridlines** 属性可返回 **GridLines** 对象，该对象代表次要网格线。可以同时返回主要网格线和次要网格线。

## 示例

```javascript
/*下例打开图表工作表 Chart1 上分类轴的主要网格线，并将网格线设置为蓝色虚线。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasMajorGridlines = true
    axis.MajorGridlines.Border.Color = RGB(0, 0, 255)
    axis.MajorGridlines.Border.LineStyle = xlDash
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表数值轴的主要网格线名称赋值给B6单元格。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    Range("B6").Value2 = axis.MajorGridlines.Name
}
```
