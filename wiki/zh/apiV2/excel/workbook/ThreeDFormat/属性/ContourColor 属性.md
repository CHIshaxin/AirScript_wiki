---
title: "ContourColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/ContourColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: COBlcOGy
---

# ContourColor (属性)

返回 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的轮廓线颜色。只读 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 类型。

## 示例

```javascript
/*此示例将第一张工作表的第一个形状三维轮廓线颜色设置为红色。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.ContourColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*本示例显示活动工作表中第二个形状三维轮廓线颜色的亮度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ThreeD.ContourColor.Brightness)
}
```
