---
title: "ShapeNodes (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/ShapeNodes%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CvUJY4UE
---

# ShapeNodes (对象)

指定的任意多边形中所有 **ShapeNode** 对象的集合。

## 说明

每一个 **ShapeNode** 对象都代表任意多边形中线段间的结点或任意多边形曲线段的控点。您可以手动创建或通过使用 **BuildFreeform** 和 **ConvertToShape** 方法来创建任意多边形。

使用 **Nodes** 属性可返回 **[ShapeNodes](./../ShapeNodes/ShapeNodes 对象.html)** 集合。

## 示例

```javascript
/*本示例删除第一张工作表上形状三中的结点四。*/
function test() {
    let shapes = Worksheets.Item(1).Shapes
    shapes.Item(3).Nodes.Delete(4)
}
```

使用 **Insert** 方法可创建一个新结点并将它添加到 **[ShapeNodes](./../ShapeNodes/ShapeNodes 对象.html)** 集合。

```javascript
/*本示例在第一张工作表上的形状三中的结点四之后添加一带有曲线段的平滑结点。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(3).Nodes
    nodes.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

使用 **Nodes**(index)（其中index是结点索引号）可返回一个 **ShapeNode** 对象。

```javascript
/*本示例将第一张工作表上形状三中的结点一设置为平滑顶点。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(3)
    shape.Nodes.SetEditingType(1, msoEditingSmooth)
}
```
