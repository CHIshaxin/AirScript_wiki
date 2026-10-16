---
title: "EndConnectionSite (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E5%B1%9E%E6%80%A7/EndConnectionSite%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ZTvZJwDy
---

# EndConnectionSite (属性)

返回一个整数，该整数指定连接符终点的连接结点。**Long** 类型，只读。

## 说明

如果指定的连接符的终点没有连接到一个形状上，则该属性产生一个错误。

## 示例

```javascript
/*此示例假定第一张工作表已包含由名为 Conn1To2 的连接符附加的两个形状。本示例的代码将向第一张工作表添加一个矩形和一个连接符。新连接符终点的连接位置与原来名为 Conn1To2 的连接符终点的连接位置相同，而新连接符的起点会连接到新矩形的第一个连接位置上。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shapes = worksheet.Shapes
    let shape = shapes.AddShape(msoShapeRectangle, 100, 420, 200, 100)
    let connFormat1 = shapes.Item("Conn1To2").ConnectorFormat
    let endConnSite1 = connFormat1.EndConnectionSite
    let endConnShape1 = connFormat1.EndConnectedShape
    let connFormat2 = shapes.AddConnector(msoConnectorCurve, 0, 0, 10, 10).ConnectorFormat
    connFormat2.BeginConnect(shape, 1)
    connFormat2.EndConnect(endConnShape1, endConnSite1)
}
```

```javascript
/*本示例显示活动工作表中第三个形状（连接符）的终点所连接的连接结点的位置。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(3)
    console.log(shape.ConnectorFormat.EndConnectionSite)
}
```
