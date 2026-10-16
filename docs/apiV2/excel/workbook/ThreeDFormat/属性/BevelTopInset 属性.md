---
title: "BevelTopInset 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/BevelTopInset%20%E5%B1%9E%E6%80%A7.html
---
# BevelTopInset (属性)

返回或设置一个值，该值指示是否应为 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象抬高顶部插入棱台，用于 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象，可读/写 **Single** 类型。

## 示例

```javascript
/*此示例设置第一张工作表的第一个形状ThreeDFormat的BevelTopInset属性。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.BevelTopInset = 20
}
```

```javascript
/*本示例显示活动工作表中第三个形状ThreeDFormat的BevelTopInset属性值。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.BevelTopInset)
}
```
