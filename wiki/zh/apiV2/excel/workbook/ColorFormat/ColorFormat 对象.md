---
title: "ColorFormat (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/ColorFormat%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: rdr21-LH
---

# ColorFormat (对象)

代表单色对象的颜色、带有渐变或图案填充的对象的前景或背景色，或者指针的颜色。

## 说明

可以将颜色设为显式的红-绿-蓝值（使用 **RGB** 属性），或设为配色方案中的一种颜色（使用 **SchemeColor** 属性）。

使用下表中列出的属性之一可返回 **[ColorFormat](./../ColorFormat/ColorFormat 对象.html)** 对象。

| 使用此属性 | 对象 | 返回一个 ColorFormat 对象，该对象代表 |
|---|---|---|
| **BackColor** | **[FillFormat](./../FillFormat/FillFormat 对象.html)** | 背景填充色（用于阴影或图案填充格式） |
| **ForeColor** | **[FillFormat](./../FillFormat/FillFormat 对象.html)** | 前景填充色（对于纯色填充格式，即代表填充颜色） |
| **BackColor** | **[LineFormat](./../LineFormat/LineFormat 对象.html)** | 线条背景色（用于图案线条） |
| **ForeColor** | **[LineFormat](./../LineFormat/LineFormat 对象.html)** | 线条前景色（对于纯色线条，即代表线条颜色） |
| **ForeColor** | **[ShadowFormat](./../ShadowFormat/ShadowFormat 对象.html)** | 阴影颜色 |
| **ExtrusionColor** | **[ThreeDFormat](./../ThreeDFormat/ThreeDFormat 对象.html)** | 有延伸的对象的侧边颜色 |

使用 **RGB** 属性可将颜色设置为显示的红-绿-蓝值。

## 示例

```javascript
/*下例向第一张工作表中添加一个矩形，然后设置矩形填充的前景色、背景色和渐变。*/
function test() {
    let fillFormat = Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    fillFormat.ForeColor.RGB = RGB(128, 0, 0)
    fillFormat.BackColor.RGB = RGB(170, 170, 170)
    fillFormat.TwoColorGradient(msoGradientHorizontal, 1)
}
```

```javascript
/*本示例显示活动工作表上第一个形状阴影的前景色的颜色类型是否为 msoColorTypeRGB。*/
function test() {
    console.log(ActiveSheet.Shapes(1).Shadow.ForeColor.Type == msoColorTypeRGB)
}
```
