---
title: "SetSegmentType (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/%E6%96%B9%E6%B3%95/SetSegmentType%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D8pkV-bF
---

# SetSegmentType (方法)

设置 *Index* 指定的结点后面的段的段类型。如果该结点为一条曲线段的控制点，则该方法设置该曲线的段类型。请注意，该方法可能由于插入或删除相邻的结点而影响结点总数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int | 必填 | 要设置段类型的顶点。 |
| SegmentType | Kso.MsoSegmentType | 必填 | 指定是直线段还是曲线段。 |

## 示例

```javascript
/*本示例将第一张工作表中第三个形状的所有线段更改为曲线段。 第三个形状必须是任意多边形。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(3).Nodes
    for (let i = 1; i <= nodes.Count; i++) {
        nodes.SetSegmentType(i, msoSegmentCurve)
    }
}
```

```javascript
/*本示例将活动工作表中第二个形状的第二个节点的线段类型设置为直线。*/
function test() {
    let nodes = ActiveSheet.Shapes.Item(2).Nodes
    nodes.SetSegmentType(2, msoSegmentLine)
}
```
