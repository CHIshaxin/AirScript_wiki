---
title: "RotationY (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/RotationY%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CYi48vpr
---

# RotationY (属性)

以度为单位返回或设置具有延伸的形状绕 y 轴的转角。取值范围在 -90 到 90 之间。正值表示向左旋转，负值表示向右旋转。**Single** 类型，可读/写。

## 说明

要设置具有延伸的形状绕 x 轴的旋转量，请使用 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的 **RotationX** 属性。要设置具有延伸的形状绕 z 轴的旋转量，请使用 **[Shape](./../../Shape/Shape 对象.html)** 对象的 **Rotation** 属性。要更改延伸的弯路径的方向而不旋转延伸的正面，请使用 **SetExtrusionDirection** 方法。

## 示例

```javascript
/*本示例向第一张工作表中添加三个相同的具有延伸的椭圆，并分别将这三个椭圆绕 y 轴的转角设置为 -30 度、0 度和 30 度。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shapes = worksheet.Shapes
    let threeDFormat1 = shapes.AddShape(msoShapeOval, 30, 30, 50, 25).ThreeD
    threeDFormat1.Visible = msoTrue
    threeDFormat1.RotationY = -30
    let threeDFormat2 = shapes.AddShape(msoShapeOval, 30, 70, 50, 25).ThreeD
    threeDFormat2.Visible = msoTrue
    threeDFormat2.RotationY = 0
    let threeDFormat3 = shapes.AddShape(msoShapeOval, 30, 110, 50, 25).ThreeD
    threeDFormat3.Visible = msoTrue
    threeDFormat3.RotationY = 30
}
```

```javascript
/*本示例显示活动工作表中第二个延伸形状绕 y 轴旋转的度数。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.RotationY)
}
```
