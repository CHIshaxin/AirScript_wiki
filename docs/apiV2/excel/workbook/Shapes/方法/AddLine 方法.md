---
title: "AddLine 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/AddLine%20%E6%96%B9%E6%B3%95.html
---
# AddLine (方法)

当本方法应用于 **[Shapes](./../../Shapes/Shapes 对象.html)** 对象时，返回一个 **[Shape](./../../Shape/Shape 对象.html)** 对象，该对象表示工作表中的新线条。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| BeginX | float | 必填 | 线条的起点相对于文档左上角的位置（以磅为单位）。 |
| BeginY | float | 必填 | 线条的起点相对于文档左上角的位置（以磅为单位）。 |
| EndX | float | 必填 | 线条的终点相对于文档左上角的位置（以磅为单位）。 |
| EndY | float | 必填 | 线条的终点相对于文档左上角的位置（以磅为单位）。 |

## 返回值

Shape

## 示例

```javascript
/*本示例向第一张工作表添加一条蓝色虚线。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let line = shapes.AddLine(10, 10, 250, 250).Line
    line.DashStyle = msoLineDashDotDot
    line.ForeColor.RGB = RGB(50, 0, 128)
}
```

```javascript
/*本示例向活动工作表添加一条带箭头的直线。*/
function test() {
    ActiveSheet.Shapes.AddLine(10, 10, 100, 100).Line.EndArrowheadStyle = msoArrowheadOpen
}
```
