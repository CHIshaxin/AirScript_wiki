---
title: "PresetGradientType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/PresetGradientType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: qY9ueT34
---

# PresetGradientType (属性)

返回指定填充的预设渐变类型。**MsoPresetGradientType**，只读。

## 说明

使用 **PresetGradient**方法可设置填充的预设渐变类型。

## 示例

```javascript
/*本示例以第一个形状中使用的填充样式对第二个形状的填充格式进行设置。*/
function test() {
    let fillFormat1 = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    if (fillFormat1.Type == msoFillGradient) {
        let fillFormat2 = Application.Worksheets.Item(1).Shapes.Item(2).Fill
        fillFormat2.Visible = true
        fillFormat2.PresetGradient(fillFormat1.GradientStyle, fillFormat1.GradientVariant, fillFormat1.PresetGradientType)
    }
}
```

```javascript
/*本示例在活动工作表中添加一个矩形，并显示填充的预设渐变类型是否为混合渐变。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
    console.log(fillFormat.PresetGradientType == msoPresetGradientMixed)
}
```
