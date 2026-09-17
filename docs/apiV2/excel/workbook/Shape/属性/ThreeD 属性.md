---
title: "ThreeD 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ThreeD%20%E5%B1%9E%E6%80%A7.html
---
# ThreeD (属性)

返回一个 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象，它包含指定形状的三维效果格式属性。只读。

## 示例

```javascript
/*本示例设置应用于活动工作表中第一个形状的3D效果的深度、延伸颜色、延伸方向以及照明方向。*/
function test() {
    let threeD = ActiveSheet.Shapes.Item(1).ThreeD
    threeD.Visible = msoTrue
    threeD.Depth = 50
    threeD.ExtrusionColor.RGB = RGB(255, 100, 255)
    //RGB value for purple
    threeD.SetExtrusionDirection(msoExtrusionTop)
    threeD.PresetLightingDirection = msoLightingLeft
}
```

```javascript
/*本示例将第一张工作表中第二个形状的3D效果的顶部棱台类型设置为msoBevelRelaxedInset。*/
function test() {
    let threeD = Worksheets.Item(1).Shapes.Item(2).ThreeD
    threeD.Visible = msoTrue
    threeD.BevelTopType = msoBevelRelaxedInset
    threeD.BevelTopDepth = 10
}
```
