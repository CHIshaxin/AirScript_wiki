---
title: "RerouteConnections (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/RerouteConnections%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BSfmw6ah
---

# RerouteConnections (方法)

此方法将重排连接在指定形状上的所有连接符；如果指定的形状是连接符，就重排该连接符。

## 说明

重排连接符使其以最短的路径连接形状。重排时，**RerouteConnections** 方法可能会断开连接符的两端并将其重新连接到形状的其他位置。

如果该方法应用于一个连接符，则只重排该连接符；如果该方法应用于一个已连接的形状，则重排该形状上所有的连接符。

请注意，**RerouteConnections** 方法用于调整连接符的大小和位置并决定要连接到哪个连接结点，因此，最初为 **BeginConnect** 和 **EndConnect** 方法的 *ConnectionSite* 参数指定的值是不相关的。

## 示例

```javascript
/*本示例将两个矩形添加到活动工作表，用曲线连接符连接两个矩形，然后重排连接符使两个矩形间采用最短的路径。*/
function test() {
    let shapes = Application.ActiveSheet.Shapes
    let firstRect = shapes.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = shapes.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let newConnector = shapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
    newConnector.ConnectorFormat.BeginConnect(firstRect, 1)
    newConnector.ConnectorFormat.EndConnect(secondRect, 1)
    newConnector.RerouteConnections()
}
```

```javascript
/*本示例将平行四边形和六边形添加到第一张工作表，用直线连接符连接两个形状，然后重排连接符使两个形状间采用最短的路径。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape1 = shapes.AddShape(msoShapeParallelogram, 100, 50, 100, 80)
    let shape2 = shapes.AddShape(msoShapeHexagon, 200, 300, 100, 80)
    let newConnector = shapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100)
    newConnector.ConnectorFormat.Type = msoConnectorStraight
    newConnector.ConnectorFormat.BeginConnect(shape1, 2)
    newConnector.ConnectorFormat.EndConnect(shape2, 3)
    newConnector.RerouteConnections()
}
```
