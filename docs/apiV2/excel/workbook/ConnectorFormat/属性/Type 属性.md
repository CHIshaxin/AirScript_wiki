---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConnectorFormat/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回或设置一个 **MsoConnectorType** 值，它代表连接符格式类型。

## 示例

```javascript
/*本示例将第一张工作表中第三个形状（连接符）格式类型设置为直线。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape = shapes.Item(3)
    shape.ConnectorFormat.Type = msoConnectorStraight
}
```

```javascript
/*本示例显示活动工作表中第四个形状（连接符）的格式类型是否为曲线。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(4)
    console.log(shape.ConnectorFormat.Type == msoConnectorCurve)
}
```
