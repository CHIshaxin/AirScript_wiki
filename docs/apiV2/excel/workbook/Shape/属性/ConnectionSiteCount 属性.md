---
title: "ConnectionSiteCount 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ConnectionSiteCount%20%E5%B1%9E%E6%80%A7.html
---
# ConnectionSiteCount (属性)

返回指定形状中的连接部位的数量。**Long** 型，只读。

## 示例

```javascript
/*本示例向活动工作表中添加两个矩形，并且用两个连接符连接这两个矩形。两个连接符的起点都连接到第一个矩形的第一个连接位置，两个连接符的终点分别连接到第二个矩形的第一个和最后一个连接位置。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let firstRect = shapes.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = shapes.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let lastsite = secondRect.ConnectionSiteCount
    let Con1 = shapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
    Con1.BeginConnect(firstRect, 1)
    Con1.EndConnect(secondRect, 1)
    let Con2 = shapes.AddConnector(msoConnectorCurve, 0, 0, 100, 100).ConnectorFormat
    Con2.BeginConnect(firstRect, 1)
    Con2.EndConnect(secondRect, lastsite)
}
```

```javascript
/*本示例首先向第一张工作表添加一个三角形，然后显示形状中的连结点的数量。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRightTriangle, 10, 10, 50, 50)
    console.log(shape.ConnectionSiteCount)
}
```
