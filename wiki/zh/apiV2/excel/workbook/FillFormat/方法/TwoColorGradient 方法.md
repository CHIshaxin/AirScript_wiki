---
title: "TwoColorGradient (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/TwoColorGradient%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BJcMEbrM
---

# TwoColorGradient (方法)

将指定填充设为双色渐变。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Style | Kso.MsoGradientStyle | 必填 | 渐变样式。 |
| Variant | int | 必填 | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 Style 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |

## 示例

```javascript
/*本示例向活动工作表中添加一个具有双色渐变填充效果的矩形，并设置填充的前景色和背景色。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
    fillFormat.ForeColor.RGB = RGB(128, 0, 0)
    fillFormat.BackColor.RGB = RGB(0, 170, 170)
    fillFormat.TwoColorGradient(msoGradientHorizontal, 1)
}
```

```javascript
/*本示例向第一张工作表中添加一个具有双色水平渐变填充效果的椭圆。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).Fill
    fillFormat.TwoColorGradient(msoGradientVertical, 2)
}
```
