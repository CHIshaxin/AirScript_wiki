---
title: "PresetLighting (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/PresetLighting%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CEq426RQ
---

# PresetLighting (属性)

返回或设置 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的延伸预设照明。可读/写 **MsoLightRigType** 类型。

## 说明

| MsoLightRigType 可以是下列 **MsoLightRigType** 常量之一。 |
| --- |
| **msoLightRigBalanced** |
| **msoLightRigBrightRoom** |
| **msoLightRigChilly** |
| **msoLightRigContrasting** |
| **msoLightRigFlat** |
| **msoLightRigFlood** |
| **msoLightRigFreezing** |
| **msoLightRigGlow** |
| **msoLightRigHarsh** |
| **msoLightRigLegacyFlat1** |
| **msoLightRigLegacyFlat2** |
| **msoLightRigLegacyFlat3** |
| **msoLightRigLegacyFlat4** |
| **msoLightRigLegacyHarsh1** |
| **msoLightRigLegacyHarsh2** |
| **msoLightRigLegacyHarsh3** |
| **msoLightRigLegacyHarsh4** |
| **msoLightRigLegacyNormal1** |
| **msoLightRigLegacyNormal2** |
| **msoLightRigLegacyNormal3** |
| **msoLightRigLegacyNormal4** |
| **msoLightRigMixed** |
| **msoLightRigMorning** |
| **msoLightRigSoft** |
| **msoLightRigSunrise** |
| **msoLightRigSunset** |
| **msoLightRigThreePoint** |
| **msoLightRigTwoPoint** |

## 示例

```javascript
/*此示例为第一张工作表的第一个形状延伸预设照明指定 Balanced 效果。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.PresetLighting = msoLightRigBalanced
}
```

```javascript
/*本示例显示活动工作表中第二个形状延伸预设照明效果是否为msoLightRigChilly。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.PresetLighting == msoLightRigChilly)
}
```
