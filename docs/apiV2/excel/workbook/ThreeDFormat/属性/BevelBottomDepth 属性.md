---
title: "BevelBottomDepth 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/BevelBottomDepth%20%E5%B1%9E%E6%80%A7.html
---
# BevelBottomDepth (属性)

返回或设置对 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象使用棱台效果时的台底深度。可读/写 **Single** 类型。

## 示例

```javascript
/*此示例对第一张工作表的第一个形状使用棱台效果时的台底深度设置为30磅。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(1)
    shape.ThreeD.Visible = msoTrue
    shape.ThreeD.BevelBottomDepth = 30
}
```

```javascript
/*本示例显示活动工作表中第二个形状使用棱台效果时的台底深度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(2)
    console.log(shape.ThreeD.BevelBottomDepth)
}
```
