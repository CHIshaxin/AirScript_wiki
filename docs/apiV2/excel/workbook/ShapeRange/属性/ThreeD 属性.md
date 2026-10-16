---
title: "ThreeD 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/ThreeD%20%E5%B1%9E%E6%80%A7.html
---
# ThreeD (属性)

返回一个 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象，它包含指定形状的三维效果格式属性。只读。

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，然后设置对象的3D效果的深度、延伸颜色、延伸方向以及照明方向。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    let threeD = shpRange.ThreeD
    threeD.Visible = true
    threeD.Depth = 50
    threeD.ExtrusionColor.RGB = RGB(255, 100, 255)
    //RGB value for purple
    threeD.SetExtrusionDirection(msoExtrusionTop)
    threeD.PresetLightingDirection = msoLightingLeft
}
```

```javascript
/*本示例将活动工作表的所有形状的3D效果的顶部棱台类型设置为msoBevelRelaxedInset。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.ThreeD.Visible = msoTrue
    Selection.ShapeRange.ThreeD.BevelTopType = msoBevelRelaxedInset
    Selection.ShapeRange.ThreeD.BevelTopDepth = 10
}
```
