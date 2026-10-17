---
title: "BevelBottomInset (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/BevelBottomInset%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DhrF4ITM
---

# BevelBottomInset (属性)

返回或设置一个值，该值指示是否应为 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象抬高底部插入棱台，可读/写 **Single** 类型。

## 示例

```javascript
/*此示例修改第一张工作表的第一个形状ThreeDFormat的BevelBottomInset属性。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(1)
    shape.ThreeD.BevelBottomInset = 20
}
```

```javascript
/*本示例显示活动工作表中第三个形状ThreeDFormat的BevelBottomInset属性值。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(3)
    console.log(shape.ThreeD.BevelBottomInset)
}
```
