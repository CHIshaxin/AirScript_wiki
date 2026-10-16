---
title: "SetEditingType (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/%E6%96%B9%E6%B3%95/SetEditingType%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D8PZWfh4
---

# SetEditingType (方法)

设置由 *Index* 指定的结点的编辑类型。如果该结点是一个曲线的控制点，本方法将对其设置相邻结点的编辑属性，该属性连接两个段。请注意，由于编辑类型的不同，本方法可能影响相邻结点的位置。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int | 必填 | 需要设置编辑类型的顶点。 |
| EditingType | Kso.MsoEditingType | 必填 | 顶点的编辑属性。 |

## 示例

```javascript
/*本示例将第一张工作表中第三个形状中的所有结点更改为光滑结点。第三个形状必须是任意多边形。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(3).Nodes
    for (let i = 1; i <= nodes.Count; i++) {
        nodes.SetEditingType(i, msoEditingSmooth)
    }
}
```

```javascript
/*本示例将活动工作表中第一个形状的第二个结点更改为对称的曲线结点。*/
function test() {
    let nodes = ActiveSheet.Shapes.Item(1).Nodes
    nodes.SetEditingType(2, msoEditingSymmetric)
}
```
