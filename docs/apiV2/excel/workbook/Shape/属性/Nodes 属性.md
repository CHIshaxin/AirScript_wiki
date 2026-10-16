---
title: "Nodes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Nodes%20%E5%B1%9E%E6%80%A7.html
---
# Nodes (属性)

返回一个 **[ShapeNodes](./../../ShapeNodes/ShapeNodes 对象.html)** 集合，它代表指定形状的几何描述。

## 说明

此属性适用于代表任意多边形的 **Shape** 或 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象。

## 示例

```javascript
/*本示例在活动工作表中第三个形状的第四个结点后添加一个带有一段曲线的平滑结点。第三个形状必须为至少带有四个顶点的任意多边形。*/
function test() {
    let nodes = ActiveSheet.Shapes.Item(3).Nodes
    nodes.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

```javascript
/*本示例显示第一张工作表中第一个形状的节点数量。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    console.log(shapes.Item(1).Nodes.Count)
}
```
