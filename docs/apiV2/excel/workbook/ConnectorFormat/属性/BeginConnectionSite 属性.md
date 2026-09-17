---
title: "BeginConnectionSite 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E5%B1%9E%E6%80%A7/BeginConnectionSite%20%E5%B1%9E%E6%80%A7.html
---
# BeginConnectionSite (属性)

返回一个整数，该整数指定连接符起点的连接结点。**Long** 类型，只读。

## 说明

如果指定的连接符的起点并未连接到任何形状上，那么本属性将产生错误。

## 示例

```javascript
/*此示例假定第一张工作表已包含由名为 Conn1To2 的连接符附加的两个形状。本示例的代码将向第一张工作表添加一个矩形和一个连接符。新连接符起点的连接位置与原来名为 Conn1To2 的连接符起点的连接位置相同，而新连接符的终点会连接到新矩形的第一个连接位置上。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shapes = worksheet.Shapes
    let shape = shapes.AddShape(msoShapeRectangle, 450, 190, 200, 100)
    shapes.AddConnector(msoConnectorCurve, 0, 0, 10, 10).Name = "Conn1To3"
    let connFormat1 = shapes.Item("Conn1To2").ConnectorFormat
    let beginConnSite1 = connFormat1.BeginConnectionSite
    let beginConnShape1 = connFormat1.BeginConnectedShape
    let connFormat2 = shapes.Item("Conn1To3").ConnectorFormat
    connFormat2.BeginConnect(beginConnShape1, beginConnSite1)
    connFormat2.EndConnect(shape, 1)
}
```

```javascript
/*本示例显示活动工作表中第三个形状（连接符）的起点所连接的连接结点的位置。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(3)
    console.log(shape.ConnectorFormat.BeginConnectionSite)
}
```
