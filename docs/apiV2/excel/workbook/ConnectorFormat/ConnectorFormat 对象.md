---
title: "ConnectorFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/ConnectorFormat%20%E5%AF%B9%E8%B1%A1.html
---
# ConnectorFormat (对象)

包含应用于连接符的属性和方法。

## 说明

连接符是用于连接其他两个形状的线，所连接的位置叫做连接结点。如果重新排列已连接的形状，那么连接符的几何形状将自动调整，以使重新排列的形状仍保持连接。

连接结点通常按下表所示的规则进行编号。

| 形状类型 | 连接结点标号方案 |
| --- | --- |
| 自选形状、艺术字、图片和 OLE 对象 | 连接结点从顶部开始按逆时针进行编号。 |
| 任意多边形 | 连接结点为顶点，与顶点编号相对应。 |

使用 **ConnectorFormat** 属性可返回 **[ConnectorFormat](./../ConnectorFormat/ConnectorFormat 对象.html)** 对象。使用 **BeginConnect** 和 **EndConnect**方法可将连接符的两端连到文档中的其他形状。使用 **RerouteConnections**方法可自动查找通过连接符连接的两个形状间的最短路径。使用 **Connector**属性可查看形状是否为连接符。

请注意，虽然向 **[Shapes](./../Shapes/Shapes 对象.html)** 集合添加连接符时，对其设置了大小和位置，但将连接符的起点和终点连接到其他形状时，连接符的大小和位置将会自动调整。因而，如果打算用连接符连接其他形状，那么对其设置的初始大小和位置就没有什么实际意义。同样，用连接符连接其他形状时，将指定要连接的形状上的连接结点，但将连接符连接好之后，用 **RerouteConnections**方法可能会改变连接符所连接的连接结点，使开始时选定的连接结点变得没有意义。

要算出一个复杂形状上各连接结点的编号，可以打开宏录制器并对形状进行试验操作，然后查看录下的代码；也可以创建一个形状并选中它，然后运行以下示例。这段代码将对每个连接结点进行编号并连接一个连接符。

## 示例

```javascript
/*本示例将对活动工作表所选形状的每个连接结点进行编号并连接一个连接符。*/
function test() {
    ActiveSheet.Shapes.SelectAll()
    let shape = Selection.ShapeRange.Item(1)
    let bx = shape.Left + shape.Width + 50
    let by = shape.Top + shape.Height + 50
    let count = shape.ConnectionSiteCount
    for (let j = 1; j <= count; j++) {
        let connector = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, bx, by, bx + 50, by + 50)
        connector.ConnectorFormat.EndConnect(shape, j)
        connector.ConnectorFormat.Type = msoConnectorElbow
        connector.Line.ForeColor.RGB = RGB(255, 0, 0)
        let l = connector.Left
        let t = connector.Top
        let textbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 36, 14)
        textbox.Fill.Visible = false
        textbox.Line.Visible = false
        textbox.TextFrame.Characters().Text = j
    }
}
```

```javascript
/*下例向第一张工作表中添加两个矩形并且用曲线连接符连接矩形。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let s = worksheet.Shapes
    let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```
