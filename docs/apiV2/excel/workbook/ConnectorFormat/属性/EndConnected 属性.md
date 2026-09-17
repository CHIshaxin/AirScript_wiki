---
title: "EndConnected 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E5%B1%9E%E6%80%A7/EndConnected%20%E5%B1%9E%E6%80%A7.html
---
# EndConnected (属性)

如果该属性值为 **msoTrue**，则指定的连接符的终点已连接到了某一形状上。**MsoTriState** 类型，只读。

## 说明

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| --- |
| **msoCTrue** 不应用于该属性。 |
| **msoFalse** 指定连接符的端点未连接到某一形状上。 |
| **msoTriStateMixed** 不应用于该属性。 |
| **msoTriStateToggle**不应用于该属性。 |
| **msoTrue** 指定连接符的端点连接到形状上。 |

## 示例

```javascript
/*如果第一张工作表上由第三个形状表示的连接符的末端连接到一个形状，则本示例将连接站点号存储在变量 oldEndConnSite 中，将对连接形状的引用存储在对象变量 oldEndConnShape 中，然后断开连接符的末端与形状的连接。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let shape = worksheet.Shapes.Item(3)
    if (shape.Connector) {
        let connectorFormat = shape.ConnectorFormat
        if (connectorFormat.EndConnected) {
            let oldEndConnSite = connectorFormat.EndConnectionSite
            let oldEndConnShape = connectorFormat.EndConnectedShape
            connectorFormat.EndDisconnect()
        }
    }
}
```

```javascript
/*本示例显示活动工作表中第四个形状（连接符）的终点是否连接到了某一形状上。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(4)
    console.log(shape.ConnectorFormat.EndConnected == msoTrue)
}
```
