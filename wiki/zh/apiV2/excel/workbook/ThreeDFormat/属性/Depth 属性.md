---
title: "Depth (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/Depth%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Ch5saXmk
---

# Depth (属性)

返回或设置一个 **Single** 值，它代表形状突出的深度。

## 说明

此属性值可以是从 -600 到 9600 的值（正值产生的延伸正面为原始形状；负值产生的延伸背面为原始形状）。

## 示例

```javascript
/*此示例向第一张工作表中添加一个椭圆，然后指定椭圆的深度为 50 磅，并指定深度部分的颜色为紫色。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.AddShape(msoShapeOval, 90, 90, 90, 40)
    let threeDFormat = shape.ThreeD
    threeDFormat.Visible = true
    threeDFormat.Depth = 50
    //RGB value for purple
    threeDFormat.ExtrusionColor.RGB = RGB(255, 100, 255)
}
```

```javascript
/*本示例显示活动工作表中第二个形状突出的深度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(2)
    console.log(shape.ThreeD.Depth)
}
```
