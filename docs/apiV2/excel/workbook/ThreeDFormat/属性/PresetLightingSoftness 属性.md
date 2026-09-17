---
title: "PresetLightingSoftness 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/PresetLightingSoftness%20%E5%B1%9E%E6%80%A7.html
---
# PresetLightingSoftness (属性)

返回或设置延伸光源效果的亮度。可读写。**MsoPresetLightingSoftness** 类型。

## 说明

| **MsoPresetLightingSoftness** 可以是下列 **MsoPresetLightingSoftness** 常量之一。 |
| --- |
| **msoLightingBright** |
| **msoLightingDim** |
| **msoLightingNormal** |
| **msoPresetLightingSoftnessMixed** |

## 示例

```javascript
/*本示例将从左侧以“强光”照亮第一张工作表中形状 1 的延伸部分。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.Visible = msoTrue
    threeDFormat.PresetLightingSoftness = msoLightingBright
    threeDFormat.PresetLightingDirection = msoLightingLeft
}
```

```javascript
/*本示例显示活动工作表中第二个形状延伸光源的亮度是否为暗光。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.PresetLightingSoftness == msoLightingDim)
}
```
