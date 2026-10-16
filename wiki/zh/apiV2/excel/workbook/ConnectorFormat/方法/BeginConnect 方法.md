---
title: "BeginConnect (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E6%96%B9%E6%B3%95/BeginConnect%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: NIBafyFU
---

# BeginConnect (方法)

将指定的连接符的起点连接到指定的形状上。如果在连接符的起点与其他形状之间已经有了连接，那么该已有的连接将中断。如果连接符的起点不在所需的连接结点上，本方法将把连接符的起点移到该连接结点，并对连接符的大小和位置作相应的调整。可用 **EndConnect** 方法将连接符的终点连接到某一形状上。

## 说明

将连接符连接到某个对象以后，该连接符的大小和位置将在必要时进行自动调整。

请注意，对 **RerouteConnections**方法的调用使得在 **BeginConnect** 方法和 **EndConnect** 方法中所指定的 *ConnectionSite* 参数值不相关联。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| ConnectedShape | Shape | 必填 | 要连接到连接符起点的形状。指定的 Shape 对象必须与连接符在同一个 Shapes 集合中。 |
| ConnectionSite | int4 | 必填 | 由 ConnectedShape 指定的形状上的连接结点。必须是一个介于 1 和指定形状的 ConnectionSiteCount 属性的返回值之间的整数。如果希望连接符自动查找它所连接的两个形状间的最短路径，请为此参数指定任何一个有效整数，并且在连接符的两端连接到形状之后使用 RerouteConnections 方法。 |

## 示例

```javascript
/*本示例将两个矩形添加到第一张工作表，并使用曲线连接线连接它们。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let s = worksheet.Shapes
    let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let c = s.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```

```javascript
/*本示例将活动工作表中第三个形状（连接线）的起点连接到第一个形状的第三个连接点，终点连接到第二个形状的第一个连接点。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape1 = shapes.Item(1)
    let shape2 = shapes.Item(2)
    let shape3 = shapes.Item(3)
    shape3.ConnectorFormat.BeginConnect(shape1, 3)
    shape3.ConnectorFormat.EndConnect(shape2, 1)
}
```
