---
title: "ConnectorFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ConnectorFormat%20%E5%B1%9E%E6%80%A7.html
---
# ConnectorFormat (属性)

返回一个包含连接符格式属性的 **[ConnectorFormat](./../../ConnectorFormat/ConnectorFormat 对象.html)** 对象。应用于代表连接符的 **[Shape](./../../Shape/Shape 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例向活动工作表中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let firstRect = shapes.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = shapes.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let c = shapes.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    let Con = c.ConnectorFormat
    Con.BeginConnect(firstRect, 1)
    Con.EndConnect(secondRect, 1)
    c.RerouteConnections()
    Con.BeginDisconnect()
    Con.EndDisconnect()
}
```

```javascript
/*本示例将第一张工作表中两个连接符格式设置为直线。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let Con1 = shapes.Item(3).ConnectorFormat
    let Con2 = shapes.Item(4).ConnectorFormat
    Con1.Type = msoConnectorStraight
    Con2.Type = msoConnectorStraight
}
```
