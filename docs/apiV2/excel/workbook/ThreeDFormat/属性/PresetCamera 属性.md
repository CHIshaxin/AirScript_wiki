---
title: "PresetCamera 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/PresetCamera%20%E5%B1%9E%E6%80%A7.html
---
# PresetCamera (属性)

返回 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的延伸预设相机。只读 **MsoPresetCamera** 类型。

## 说明

| MsoPresetCamera 可以是下列 **MsoPresetCamera** 常量之一。 |
| --- |
| **msoCameraIsometricBottomDown** |
| **msoCameraIsometricBottomUp** |
| **msoCameraIsometricLeftDown** |
| **msoCameraIsometricLeftUp** |
| **msoCameraIsometricOffAxis1Left** |
| **msoCameraIsometricOffAxis1Right** |
| **msoCameraIsometricOffAxis1Top** |
| **msoCameraIsometricOffAxis2Left** |
| **msoCameraIsometricOffAxis2Right** |
| **msoCameraIsometricOffAxis2Top** |
| **msoCameraIsometricOffAxis3Bottom** |
| **msoCameraIsometricOffAxis3Left** |
| **msoCameraIsometricOffAxis3Right** |
| **msoCameraIsometricOffAxis4Bottom** |
| **msoCameraIsometricOffAxis4Left** |
| **msoCameraIsometricOffAxis4Right** |
| **msoCameraIsometricRightDown** |
| **msoCameraIsometricRightUp** |
| **msoCameraIsometricTopDown** |
| **msoCameraIsometricTopUp** |
| **msoCameraLegacyObliqueBottom** |
| **msoCameraLegacyObliqueBottomLeft** |
| **msoCameraLegacyObliqueBottomRight** |
| **msoCameraLegacyObliqueFront** |
| **msoCameraLegacyObliqueLeft** |
| **msoCameraLegacyObliqueRight** |
| **msoCameraLegacyObliqueTop** |
| **msoCameraLegacyObliqueTopLeft** |
| **msoCameraLegacyObliqueTopRight** |
| **msoCameraLegacyPerspectiveBottom** |
| **msoCameraLegacyPerspectiveBottomLeft** |
| **msoCameraLegacyPerspectiveBottomRight** |
| **msoCameraLegacyPerspectiveFront** |
| **msoCameraLegacyPerspectiveLeft** |
| **msoCameraLegacyPerspectiveRight** |
| **msoCameraLegacyPerspectiveTop** |
| **msoCameraLegacyPerspectiveTopLeft** |
| **msoCameraLegacyPerspectiveTopRight** |
| **msoCameraObliqueBottom** |
| **msoCameraObliqueBottomLeft** |
| **msoCameraObliqueBottomRight** |
| **msoCameraObliqueLeft** |
| **msoCameraObliqueRight** |
| **msoCameraObliqueTop** |
| **msoCameraObliqueTopLeft** |
| **msoCameraObliqueTopRight** |
| **msoCameraOrthographicFront** |
| **msoCameraPerspectiveAbove** |
| **msoCameraPerspectiveAboveLeftFacing** |
| **msoCameraPerspectiveAboveRightFacing** |
| **msoCameraPerspectiveBelow** |
| **msoCameraPerspectiveContrastingLeftFacing** |
| **msoCameraPerspectiveContrastingRightFacing** |
| **msoCameraPerspectiveFront** |
| **msoCameraPerspectiveHeroicExtremeLeftFacing** |
| **msoCameraPerspectiveHeroicExtremeRightFacing** |
| **msoCameraPerspectiveHeroicLeftFacing** |
| **msoCameraPerspectiveHeroicRightFacing** |
| **msoCameraPerspectiveLeft** |
| **msoCameraPerspectiveRelaxed** |
| **msoCameraPerspectiveRelaxedModerately** |
| **msoCameraPerspectiveRight** |
| **msoPresetCameraMixed** |

## 示例

```javascript
/*此示例显示第一张工作表的第一个形状 ThreeDFormat 对象的延伸预设相机是否为msoCameraLegacyObliqueTopRight。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    console.log(shape.ThreeD.PresetCamera == msoCameraLegacyObliqueTopRight)
}
```

```javascript
/*本示例显示活动工作表中第二个形状 ThreeDFormat 对象的延伸预设相机是否为msoCameraIsometricLeftUp。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.PresetCamera == msoCameraIsometricLeftUp)
}
```
