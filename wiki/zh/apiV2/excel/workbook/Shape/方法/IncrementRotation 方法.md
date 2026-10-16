---
title: "IncrementRotation (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/IncrementRotation%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bb-SVTRc
---

# IncrementRotation (方法)

使指定的形状按指定度数值绕 Z 轴旋转。使用 **Rotation** 属性可设置形状的绝对转角。

## 说明

如果要将三维形状绕 X 轴或 Y 轴旋转，请分别使用 **IncrementRotationX** 方法或 **IncrementRotationY** 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定形状的水平旋转量，以度为单位。正值为顺时针旋转形状，负值逆时针旋转形状。 |

## 示例

```javascript
/*本示例将活动工作表的第二个形状设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.Fill.PresetTextured(msoTextureGranite)
    shape.IncrementLeft(70)
    shape.IncrementTop(-50)
    shape.IncrementRotation(30)
}
```

```javascript
/*本示例将第一张工作表的第一个形状向上移 100 磅，然后逆时针旋转 30 度。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(1)
    shape.IncrementTop(-100)
    shape.IncrementRotation(-30)
}
```
