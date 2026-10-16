---
title: "PresetGradient 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E6%96%B9%E6%B3%95/PresetGradient%20%E6%96%B9%E6%B3%95.html
---
# PresetGradient (方法)

将指定填充设为一个预设的渐变。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Style | Kso.MsoGradientStyle | 必填 | 渐变样式。 |
| Variant | int | 必填 | 渐变变量。取值范围为 1 到 4 之间，分别与“填充效果”对话框中“渐变”选项卡上的四个渐变变量相对应。如果 GradientStyle 设为 msoGradientFromCenter，则 Variant 参数只能设为 1 或 2。 |
| PresetGradientType | Kso.MsoPresetGradientType | 必填 | 预设渐变类型。 |

## 示例

```javascript
/*本示例使用预设的渐变样式填充第一张工作表中第一个形状。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    fillFormat.PresetGradient(msoGradientHorizontal, 3, msoGradientFog)
}
```

```javascript
/*本示例在活动工作表中添加一个矩形，然后使用预设的渐变样式填充矩形。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 100, 100, 50).Fill
    fillFormat.PresetGradient(msoGradientVertical, 1, msoGradientBrass)
}
```
