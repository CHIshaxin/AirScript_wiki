---
title: "BeginDisconnect (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E6%96%B9%E6%B3%95/BeginDisconnect%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: sV4eDDuo
---

# BeginDisconnect (方法)

使指定的连接符的起点与其所连接的形状脱离。本方法并不修改连接符的尺寸和位置：连接符的起点仍保留在原来所连接的连接结点的位置，但与该连接结点之间不再有连接。可用 **EndDisconnect** 方法使连接符的终点与某一形状脱离。

## 示例

```javascript
/*本示例向第一张工作表中添加两个矩形，用一个连接符连接这两个矩形，并自动将连接符路径修改为最短路径，然后断开矩形间的连接符。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let s = worksheet.Shapes
    let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
    c.ConnectorFormat.BeginDisconnect()
    c.ConnectorFormat.EndDisconnect()
}
```

```javascript
/*本示例使活动工作表中第三个形状（连接线）的起点与其所连接的形状脱离*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape3 = shapes.Item(3)
    shape3.ConnectorFormat.BeginDisconnect()
}
```
