---
title: "PresetShape 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/PresetShape%20%E5%B1%9E%E6%80%A7.html
---
# PresetShape (属性)

返回或设置指定艺术字的形状样式。**MsoPresetTextEffectShape** 类型，可读写。

## 说明

| **MsoPresetTextEffectShape**可以是下列 **MsoPresetTextEffectShape** 常量之一。 |
| --- |
| **msoTextEffectShapeArchDownCurve** |
| **msoTextEffectShapeArchDownPour** |
| **msoTextEffectShapeArchUpCurve** |
| **msoTextEffectShapeArchUpPour** |
| **msoTextEffectShapeButtonCurve** |
| **msoTextEffectShapeButtonPour** |
| **msoTextEffectShapeCanDown** |
| **msoTextEffectShapeCanUp** |
| **msoTextEffectShapeCascadeDown** |
| **msoTextEffectShapeCascadeUp** |
| **msoTextEffectShapeChevronDown** |
| **msoTextEffectShapeChevronUp** |
| **msoTextEffectShapeCircleCurve** |
| **msoTextEffectShapeCirclePour** |
| **msoTextEffectShapeCurveDown** |
| **msoTextEffectShapeCurveUp** |
| **msoTextEffectShapeDeflate** |
| **msoTextEffectShapeDeflateBottom** |
| **msoTextEffectShapeDeflateInflateDeflate** |
| **msoTextEffectShapeDoubleWave1** |
| **msoTextEffectShapeFadeDown** |
| **msoTextEffectShapeFadeRight** |
| **msoTextEffectShapeInflate** |
| **msoTextEffectShapeInflateTop** |
| **msoTextEffectShapePlainText** |
| **msoTextEffectShapeRingOutside** |
| **msoTextEffectShapeSlantUp** |
| **msoTextEffectShapeTriangleDown** |
| **msoTextEffectShapeWave1** |
| **msoTextEffectShapeDeflateInflate** |
| **msoTextEffectShapeDeflateTop** |
| **msoTextEffectShapeDoubleWave2** |
| **msoTextEffectShapeFadeLeft** |
| **msoTextEffectShapeFadeUp** |
| **msoTextEffectShapeInflateBottom** |
| **msoTextEffectShapeMixed** |
| **msoTextEffectShapeRingInside** |
| **msoTextEffectShapeSlantDown** |
| **msoTextEffectShapeStop** |
| **msoTextEffectShapeTriangleUp** |
| **msoTextEffectShapeWave2** |

设置 **PresetTextEffect** 属性会自动设置 **PresetShape** 属性。

## 示例

```javascript
/*本示例将第一张工作表上所有艺术字的形状设置为中心点向下的 V 形。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    for (let s = 1; s <= worksheet.Shapes.Count; s++) {
        if (worksheet.Shapes.Item(s).Type == msoTextEffect) {
            worksheet.Shapes.Item(s).TextEffect.PresetShape = msoTextEffectShapeChevronDown
        }
    }
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字的形状是否成功设置成msoTextEffectShapeChevronUp。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.TextEffect.PresetShape = msoTextEffectShapeChevronUp
    console.log(shape.TextEffect.PresetShape == msoTextEffectShapeChevronUp)
}
```
