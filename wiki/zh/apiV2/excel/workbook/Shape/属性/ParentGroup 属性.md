---
title: "ParentGroup (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ParentGroup%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D6_ytVyM
---

# ParentGroup (属性)

返回 **[Shape](./../../Shape/Shape 对象.html)** 对象，它代表子形状或子形状区域的共同父形状。

## 示例

```javascript
/*本示例向活动工作表添加两个形状，然后通过删除组合的父形状来删除这两个形状。*/
function test() {
    let shapes = Application.ActiveSheet.Shapes
    shapes.AddShape(1, 10, 10, 100, 100)
    shapes.AddShape(2, 110, 120, 100, 100)
    shapes.Range([1, 2]).Group()
    // Using the child shape in the group get the Parent shape.
    let pgShape = ActiveSheet.Shapes.Item(1).GroupItems.Item(1).ParentGroup
    console.log("The two shapes will now be deleted.")
    // Delete the parent shape.
    pgShape.Delete()
}
```

```javascript
/*本示例向第一张工作表添加两个形状，然后组合在同一组，并重新命名它们的父形状。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.AddShape(1, 10, 10, 100, 100)
    shapes.AddShape(2, 150, 10, 70, 100)
    shapes.Range([1, 2]).Group()
    let pgShape = ActiveSheet.Shapes.Item(1).GroupItems.Item(1).ParentGroup
    pgShape.Name = "父形状"
}
```
