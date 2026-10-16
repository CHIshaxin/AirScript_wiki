---
title: "SetExtrusionDirection 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E6%96%B9%E6%B3%95/SetExtrusionDirection%20%E6%96%B9%E6%B3%95.html
---
# SetExtrusionDirection (方法)

设置延伸经过的路径偏离延伸形状的方向。

## 说明

| **MsoPresetExtrusionDirection** 可以是下列 **MsoPresetExtrusionDirection** 常量之一。 |
| --- |
| **msoExtrusionBottom** |
| **msoExtrusionBottomLeft** |
| **msoExtrusionBottomRight** |
| **msoExtrusionLeft** |
| **msoExtrusionNone** |
| **msoExtrusionRight** |
| **msoExtrusionTop** |
| **msoExtrusionTopLeft** |
| **msoExtrusionTopRight** |
| **msoPresetExtrusionDirectionMixed** |

此方法将 **PresetExtrusionDirection** 属性值设置为由 *PresetExtrusionDirection* 参数指定的方向。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| PresetExtrusionDirection | Kso.MsoPresetExtrusionDirection | 必填 | 指定延伸方向。 |

## 示例

```javascript
/*本示例使第一张工作表中的第一个形状向上延伸，并使光线从左方照射。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.Visible = true
    threeDFormat.SetExtrusionDirection(msoExtrusionTop)
    threeDFormat.PresetLightingDirection = msoLightingLeft
}
```

```javascript
/*本示例将活动工作表中第三个形状的 3D 效果的延伸方向设置成右上角。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    shape.ThreeD.Visible = msoTrue
    shape.ThreeD.SetExtrusionDirection(msoExtrusionTopRight)
}
```
