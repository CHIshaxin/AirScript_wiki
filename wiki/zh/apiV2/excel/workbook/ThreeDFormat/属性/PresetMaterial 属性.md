---
title: "PresetMaterial (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/PresetMaterial%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B54D-bZz
---

# PresetMaterial (属性)

返回或设置延伸部分的表面材料。**MsoPresetMaterial**，可读写。

## 说明

| **MsoPresetMaterial** 可以是下列 **MsoPresetMaterial** 常量之一。 |
| --- |
| **msoMaterialMatte** |
| **msoMaterialMetal** |
| **msoMaterialPlastic** |
| **msoMaterialWireFrame** |
| **msoPresetMaterialMixed** |

## 示例

```javascript
/*本示例指定第一张工作表中第一个形状的延伸表面为线框。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.Visible = msoTrue
    threeDFormat.PresetMaterial = msoMaterialWireFrame
}
```

```javascript
/*本示例显示活动工作表中第二个形状延伸表面是否为金属。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.PresetMaterial == msoMaterialMetal)
}
```
