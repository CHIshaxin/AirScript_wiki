---
title: "OneColorGradient (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/OneColorGradient%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BwntshH0
---

# OneColorGradient (方法)

将指定填充设为单色渐变。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Style | Kso.MsoGradientStyle | 必填 | 渐变样式。 |
| Variant | int | 必填 | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| Degree | float | 必填 | 渐变程度。可以为 0.0（暗）到 1.0（亮）之间的值。 |

## 示例

```javascript
/*本示例将活动工作表中第一个形状设置成具有单色水平渐变填充效果。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(1).Fill
    fillFormat.OneColorGradient(msoGradientHorizontal, 1, 0.5)
}
```

```javascript
/*本示例向第一张工作表中添加一个具有单色垂直渐变填充效果的椭圆。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).Fill
    fillFormat.OneColorGradient(msoGradientVertical, 1, 0.3)
}
```
