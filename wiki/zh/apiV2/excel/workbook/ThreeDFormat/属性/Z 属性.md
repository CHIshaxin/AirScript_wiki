---
title: "Z (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/Z%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C_VsS8e4
---

# Z (属性)

返回指定 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象的 Z 次序。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状ThreeDFormat对象的 Z 次序设置成100。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let threeDFormat = worksheet.Shapes.Item(1).ThreeD
    threeDFormat.Z = 100
}
```

```javascript
/*本示例显示活动工作表中第三个形状ThreeDFormat对象的 Z 次序。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.Z)
}
```
