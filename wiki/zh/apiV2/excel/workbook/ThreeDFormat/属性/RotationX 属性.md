---
title: "RotationX (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/RotationX%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C_iJDoIQ
---

# RotationX (属性)

返回或设置延伸形状绕 x 轴旋转的度数。取值范围在 –90 到 90 之间。正值表示向上旋转；负值表示向下旋转。**Single** 类型，可读写。

## 说明

要设置延伸形状绕 y 轴的旋转，请使用 ThreeDFormat 对象的 **RotationY** 属性。要设置具有延伸的形状绕 z 轴的旋转量，请使用 **[Shape](./../../Shape/Shape 对象.html)** 对象的 **Rotation**属性。要更改延伸的弯路径的方向而不旋转延伸的正面，请使用 **SetExtrusionDirection** 方法。

## 示例

```javascript
/*本示例向第一张工作表中添加三个相同的具有延伸的椭圆，并分别将这三个椭圆绕 x 轴的转角设置为 -30 度、0 度和 30 度。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shapes = worksheet.Shapes
    let threeDFormat1 = shapes.AddShape(msoShapeOval, 30, 30, 50, 25).ThreeD
    threeDFormat1.Visible = msoTrue
    threeDFormat1.RotationX = -30
    let threeDFormat2 = shapes.AddShape(msoShapeOval, 30, 70, 50, 25).ThreeD
    threeDFormat2.Visible = msoTrue
    threeDFormat2.RotationX = 0
    let threeDFormat3 = shapes.AddShape(msoShapeOval, 30, 110, 50, 25).ThreeD
    threeDFormat3.Visible = msoTrue
    threeDFormat3.RotationX = 30
}
```

```javascript
/*本示例显示活动工作表中第二个延伸形状绕 x 轴旋转的度数。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.RotationX)
}
```
