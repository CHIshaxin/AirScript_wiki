---
title: "ExtrusionColorType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/ExtrusionColorType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D4lc29If
---

# ExtrusionColorType (属性)

返回或设置一个值，该值表示延伸颜色是否基于延伸形状的填充（延伸前表面）并在形状填充更改时自动进行更改，或延伸颜色是否独立于形状填充。**MsoExtrusionColorType** 类型，可读写。

## 说明

| **MsoExtrusionColorType** 可以是下列 **MsoExtrusionColorType** 常量之一。 |
| --- |
| **msoExtrusionColorAutomatic** 基于形状填充的延伸颜色。 |
| **msoExtrusionColorTypeMixed** |
| **msoExtrusionColorCustom** 独立于形状填充的延伸颜色。 |

## 示例

```javascript
/*如果第一张工作表中第一个形状有自动延伸色，则本示例将将延伸设置为自定义的黄色。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    if (threeDFormat.ExtrusionColorType == msoExtrusionColorAutomatic) {
        threeDFormat.ExtrusionColor.RGB = RGB(240, 235, 16)
    }
}
```

```javascript
/*本示例显示活动工作表中第二个形状的延伸颜色是否独立于形状填充。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.ExtrusionColorType == msoExtrusionColorCustom)
}
```
