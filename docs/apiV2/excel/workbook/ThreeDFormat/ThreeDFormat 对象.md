---
title: "ThreeDFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/ThreeDFormat%20%E5%AF%B9%E8%B1%A1.html
---
# ThreeDFormat (对象)

该对象代表一个形状的三维格式。

## 说明

不能对某些形状应用三维格式，例如斜截形状或多处间断的路径。对这些形状， **[ThreeDFormat](./../ThreeDFormat/ThreeDFormat 对象.html)** 对象的大多数属性和方法将失败。

使用 **ThreeD** 属性可返回一个 **[ThreeDFormat](./../ThreeDFormat/ThreeDFormat 对象.html)** 对象。

## 示例

```javascript
/*以下示例向第一张工作表添加一个椭圆，然后指定椭圆延伸至 50 磅深度，延伸为紫色。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.AddShape(msoShapeOval, 90, 90, 90, 40)
    let threeDFormat = shape.ThreeD
    threeDFormat.Visible = true
    threeDFormat.Depth = 50
    threeDFormat.ExtrusionColor.RGB = RGB(255, 100, 255)
    //RGB value for purple
}
```

```javascript
/*本示例显示活动工作表中第三个形状突出的深度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(3)
    console.log(shape.ThreeD.Depth)
}
```
