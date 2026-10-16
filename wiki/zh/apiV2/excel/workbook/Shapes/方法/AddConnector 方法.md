---
title: "AddConnector (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/AddConnector%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DKKTG4W0
---

# AddConnector (方法)

创建一个连接符。返回一个代表新连接符的 **[Shape](./../../Shape/Shape 对象.html)** 对象。添加一个连接符时，它没有连接到任何对象。使用 **BeginConnect** 和 **EndConnect** 方法可将连接符的头和尾连接到文档中的其他形状上。

## 说明

| **MsoConnectorType** 可以是下列 **MsoConnectorType** 常量之一。 |
| --- |
| **msoConnectorElbow** |
| **msoConnectorTypeMixed** |
| **msoConnectorCurve** |
| **msoConnectorStraight** |

将连接符连接到形状上时，连接符的大小和位置将自动进行必要的调整。因此，将连接符连接到其他形状上时，添加连接符时所指定的初始大小和位置是无关紧要的。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | Kso.MsoConnectorType | 必填 | 要添加的连接符类型。 |
| BeginX | float | 必填 | 连接符的起点相对于文档左上角的水平位置（以磅为单位）。 |
| BeginY | float | 必填 | 连接符的起点相对于文档左上角的垂直位置（以磅为单位）。 |
| EndX | float | 必填 | 连接符的终点相对于文档左上角的水平位置（以磅为单位）。 |
| EndY | float | 必填 | 连接符的终点相对于文档左上角的垂直位置（以磅为单位）。 |

## 返回值

Shape

## 示例

```javascript
/*本示例向第一张工作表添加直线连接线。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let connector = shapes.AddConnector(msoConnectorStraight, 150, 150, 200, 200)
}
```

```javascript
/*本示例向活动工作表添加肘形连接符。*/
function test() {
    ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 100, 100, 400, 200)
}
```
