---
title: "Line (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Line%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CXfSvITK
---

# Line (属性)

返回一个 **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象，它包含指定形状的线条格式属性。对于线条， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表线条本身；而对于带有边框的形状， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表边框。只读。

## 示例

```javascript
/*本示例向活动工作表添加一条蓝色虚线。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddLine(10, 10, 250, 250)
    shape.Line.DashStyle = msoLineDashDotDot
    shape.Line.ForeColor.RGB = RGB(50, 0, 128)
}
```

```javascript
/*本示例将十字添加到第一张工作表中，然后将其边框设置为8磅厚和红色。*/
function test() {
    let shapes = Worksheets.Item(1).Shapes
    let shape = shapes.AddShape(msoShapeCross, 10, 10, 50, 70)
    shape.Line.Weight = 8
    shape.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
