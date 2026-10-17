---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ThreeDFormat/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CrGtJb-D
---

# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*此示例将第一张工作表的第一个形状的 3D 效果设置为可见。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(1)
    shape.ThreeD.Visible =  msoTrue
}
```

```javascript
/*本示例显示活动工作表中第三个形状的 3D 效果是否可见。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.ThreeD.Visible == msoTrue)
}
```
