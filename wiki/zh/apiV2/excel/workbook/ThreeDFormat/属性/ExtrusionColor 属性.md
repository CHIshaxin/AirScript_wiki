---
title: "ExtrusionColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/ExtrusionColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D7T8ckDu
---

# ExtrusionColor (属性)

返回一个 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 对象，该对象表示形状延伸区域的颜色。只读。

## 示例

```javascript
/*本示例向第一张工作表中添加椭圆，然后指定椭圆的深度为 50 磅，并指定深度部分的颜色为紫色。*/
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
/*本示例显示活动工作表中第三个形状的延伸的颜色是否为红色。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.ExtrusionColor.RGB == RGB(255, 0, 0))
}
```
