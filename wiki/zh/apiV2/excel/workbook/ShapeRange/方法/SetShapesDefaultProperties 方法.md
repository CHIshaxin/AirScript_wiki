---
title: "SetShapesDefaultProperties (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/SetShapesDefaultProperties%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DpdQoPS1
---

# SetShapesDefaultProperties (方法)

将指定形状的格式设置为形状的默认格式。

## 示例

```javascript
/*本示例向第一张工作表中添加矩形，通过新建的shpRange对象设置该矩形的填充样式，将矩形的格式设置为形状的默认格式，然后在工作表中添加另一个较小的矩形。第二个矩形的填充格式与第一个矩形一样。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape1 = shapes.AddShape(msoShapeRectangle, 5, 5, 80, 60)
    let shpRange = shapes.Range([shape1.Name])
    let f = shpRange.Fill
    f.ForeColor.RGB = (0, 0, 255)
    f.BackColor.RGB = (0, 204, 255)
    f.Patterned(msoPatternHorizontalBrick)
    // Set formatting as default formatting
    shpRange.SetShapesDefaultProperties()
    // Create new shape with default formatting
    shapes.AddShape(msoShapeRectangle, 90, 90, 40, 30)
}
```

```javascript
/*本示例将活动工作表中新建的shpRange对象的格式设置为默认格式，然后添加一个平行四边形，新形状的格式则与shpRange对象一样。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1])
    shpRange.SetShapesDefaultProperties()
    ActiveSheet.Shapes.AddShape(msoShapeParallelogram, 200, 200, 100, 150)
}
```
