---
title: "SetShapesDefaultProperties 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/SetShapesDefaultProperties%20%E6%96%B9%E6%B3%95.html
---
# SetShapesDefaultProperties (方法)

将指定形状的格式设置为形状的默认格式。

## 示例

```javascript
/*本示例向活动工作表中添加矩形，设置该矩形的填充样式，将矩形的格式设置为该矩形的默认格式，然后在文档中添加另一个较小的矩形。第二个矩形的填充格式与第一个矩形一样。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape1 = shapes.AddShape(msoShapeRectangle, 5, 5, 80, 60)
    let f = shape1.Fill
    f.ForeColor.RGB = (0, 0, 255)
    f.BackColor.RGB = (0, 204, 255)
    f.Patterned(msoPatternHorizontalBrick)
    // Set formatting as default formatting
    shape1.SetShapesDefaultProperties()
    // Create new shape with default formatting
    shapes.AddShape(msoShapeRectangle, 90, 90, 40, 30)
}
```

```javascript
/*本示例将第一张工作表的第一个形状的格式设置为默认格式，然后添加一个平行四边形，使其填充格式与第一个形状一样。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).SetShapesDefaultProperties()
    shapes.AddShape(msoShapeParallelogram, 200, 200, 100, 150)
}
```
