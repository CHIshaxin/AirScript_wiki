---
title: "FillFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/FillFormat%20%E5%AF%B9%E8%B1%A1.html
---
# FillFormat (对象)

代表形状的填充格式。

## 说明

形状可以有纯色、渐变、纹理、图案、图片或半透明填充。

**[FillFormat](./../FillFormat/FillFormat 对象.html)** 对象的很多属性是只读的。要设置这些属性中每一个，必须使用相应的方法。

使用 **Fill**属性可返回 **[FillFormat](./../FillFormat/FillFormat 对象.html)** 对象。

## 示例

```javascript
/*下例向第一张工作表中添加矩形并且设置矩形填充的渐变和颜色。*/
function test() {
    let sheet = Application.Worksheets.Item(1)
    let shapes = sheet.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
    shapes.ForeColor.RGB = (0, 128, 128)
    shapes.OneColorGradient(msoGradientHorizontal, 1, 1)
}
```

```javascript
/*本示例在活动工作表中添加椭圆，并将前景色设置为红色，填充图案设置为深色竖线。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeOval, 0, 0, 40, 80).Fill
    fillFormat.ForeColor.RGB = RGB(255, 0, 0)
    fillFormat.Patterned(msoPatternDarkVertical)
}
```
