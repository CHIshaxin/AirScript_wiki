---
title: "FieldOfView (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/FieldOfView%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B0Gbu87S
---

# FieldOfView (属性)

返回或设置可以查看 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的角度。可读/写 **Single** 类型。

## 示例

```javascript
/*此示例设置第一张工作表的第一个形状ThreeDFormat对象的角度。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.ThreeD.FieldOfView = 20
}
```

```javascript
/*本示例显示活动工作表中第三个形状ThreeDFormat对象的角度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.FieldOfView)
}
```
