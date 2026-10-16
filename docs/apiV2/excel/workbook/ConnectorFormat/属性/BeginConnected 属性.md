---
title: "BeginConnected 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E5%B1%9E%E6%80%A7/BeginConnected%20%E5%B1%9E%E6%80%A7.html
---
# BeginConnected (属性)

如果指定的连接符的起点已连接到了某一形状上，则该属性值为 **msoTrue**。**MsoTriState** 类型，只读。

## 说明

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** |
| **msoFalse** |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue**：指定连接符的起点与形状相连。 |

## 示例

```javascript
/*如果第一张工作表上的第三个形状是一个连接符，其起点连接到形状，则本示例将连接站点号存储在变量 oldBeginConnSite中，在对象变量 oldBeginConnShape中存储对连接形状的引用，然后将连接符的开头与形状断开连接。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(3)
    if (shape.Connector) {
        let connectorFormat = shape.ConnectorFormat
        if (connectorFormat.BeginConnected) {
            let oldBeginConnSite = connectorFormat.BeginConnectionSite
            let oldBeginConnShape = connectorFormat.BeginConnectedShape
            connectorFormat.BeginDisconnect()
        }
    }
}
```

```javascript
/*本示例显示活动工作表中第四个形状（连接符）的起点是否连接到了某一形状上。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(4)
    console.log(shape.ConnectorFormat.BeginConnected == msoTrue)
}
```
