---
title: "Line (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Line%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 2GhM-EUE
---

# Line (属性)

返回一个 **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象，它包含指定形状的线条格式属性。对于线条， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表线条本身；而对于带有边框的形状， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表边框。只读。

## 示例

```javascript
/*本示例在第一张工作表添加两条线，用新建的ShapeRange对象将两条线设置为蓝色虚线。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape1 = shapes.AddLine(10, 10, 250, 250)
    let shape2 = shapes.AddLine(20, 20, 150, 250)
    let shpRange = shapes.Range([shape1.Name, shape2.Name])
    shpRange.Line.DashStyle = msoLineDashDotDot
    shpRange.Line.ForeColor.RGB = RGB(50, 0, 128)
}
```

```javascript
/*本示例将活动工作表的所有形状的边框粗细设置为 8 磅，颜色为红色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Line.Weight = 8
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
