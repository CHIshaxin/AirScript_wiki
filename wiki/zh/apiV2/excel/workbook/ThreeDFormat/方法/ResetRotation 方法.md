---
title: "ResetRotation (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E6%96%B9%E6%B3%95/ResetRotation%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: VK4RaNQV
---

# ResetRotation (方法)

将延伸部分绕 x 轴和 y 轴的转角重新设置为 0（零）度，从而使延伸部分的前面朝前。本方法并不重新设置绕 z 轴的转角。

## 说明

要将延伸部分绕 X 轴和 Y 轴的转角设置为 0（零）度之外的任意数，请使用 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的 **RotationX** 和 **RotationY** 属性。要设置延伸部分绕 Z 轴的转角，请使用表示延伸形状的 **[Shape](./../../Shape/Shape 对象.html)** 对象的 **Rotation** 属性。

## 示例

```javascript
/*此示例将第一张工作表上第一个形状的延伸部分绕 x 轴和 y 轴的转角重设为 0（零）度。*/
function test() {
    let worksheet = Worksheets.Item(1)
    worksheet.Shapes.Item(1).ThreeD.ResetRotation()
}
```

```javascript
/*本示例将活动工作表第三个形状的延伸部分绕 x 轴和 y 轴的转角重设为 0（零）度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    shape.ThreeD.ResetRotation()
}
```
