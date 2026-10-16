---
title: "IncrementRotationX 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E6%96%B9%E6%B3%95/IncrementRotationX%20%E6%96%B9%E6%B3%95.html
---
# IncrementRotationX (方法)

按指定度数更改指定形状绕 X 轴的旋转角度。可以使用 **RotationX** 属性设置形状绕 X 轴的绝对旋转角度。

## 说明

不能使指定形状绕 x 轴的转角超出 **RotationX** 属性值的上下界（90 度到 -90 度）。例如，如果 **RotationX** 属性初始值为 80，并将 *Increment* 参数指定为 40，则结果转角将为 90 度（即 **RotationX** 属性值的上界），而不是 120 度。

要更改形状绕 Y 轴的旋转度数，请使用 **IncrementRotationY** 方法；要更改图形绕 Z 轴的旋转度数，请使用 **IncrementRotationZ** 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定形状绕 X 轴旋转角度的改变量（以度为单位）。取值范围在 -90 到 90 之间。正值表示形状向上旋转，负值表示形状向下旋转。 |

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状向上倾斜 10 度。该形状必须是挤制形状才能看到本示例代码的效果。*/
function test() {
    let worksheet = Worksheets.Item(1)
    worksheet.Shapes.Item(1).ThreeD.IncrementRotationX(10)
}
```

```javascript
/*本示例将活动工作表第三个形状向下倾斜 20 度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    shape.ThreeD.IncrementRotationX(-20)
}
```
