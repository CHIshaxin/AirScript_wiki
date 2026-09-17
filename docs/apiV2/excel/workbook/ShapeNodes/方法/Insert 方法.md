---
title: "Insert 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeNodes/%E6%96%B9%E6%B3%95/Insert%20%E6%96%B9%E6%B3%95.html
---
# Insert (方法)

在任意多边形形状中插入节点。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int | 必填 | **Long** 型。形状节点的编号，将在该节点后插入新节点。 |
| SegmentType | Kso.MsoSegmentType | 必填 | 线段类型。 |
| EditingType | Kso.MsoEditingType | 必填 | 编辑的类型。 |
| X1 | float | 必填 | 如果新线段的 *EditingType* 为 **msoEditingAuto**，那么此参数就以磅为单位指定文档的左上角与新线段的终点之间的水平距离。如果新节点的 *EditingType* 为 **msoEditingCorner**，那么此参数就以磅为单位指定文档的左上角与新线段的第一个控制点之间的水平距离。 |
| Y1 | float | 必填 | 如果新线段的 *EditingType* 为 **msoEditingAuto**，则此参数就以磅为单位指定从文档左上角到新线段终点的垂直距离。如果新节点的 *EditingType* 属性为 **msoEditingCorner**，则此参数就以磅为单位指定从文档左上角到新线段第一个控制点的垂直距离。 |
| X2 | Default<float, 0.000000> | 可选 | 如果新线段的 *EditingType* 属性为 **msoEditingCorner**，则此参数就以磅为单位指定从文档左上角到新线段第二个控制点的水平距离。如果新线段的 *EditingType* 为 **msoEditingAuto**，则不用为此参数指定值。 |
| Y2 | Default<float, 0.000000> | 可选 | 如果新线段的 *EditingType* 为 **msoEditingCorner**，则此参数就以磅为单位指定从文档左上角到新线段第二个控制点的垂直距离。如果新线段的 *EditingType* 为 **msoEditingAuto**，则不用为此参数指定值。 |
| X3 | Default<float, 0.000000> | 可选 | 如果新线段的 *EditingType* 为 **msoEditingCorner**，则此参数就以磅为单位指定从文档的左上角到新线段终点之间的水平距离。如果新线段的 *EditingType* 为 **msoEditingAuto**，则不用为此参数指定值。 |
| Y3 | Default<float, 0.000000> | 可选 | 如果新线段的 *EditingType* 为 **msoEditingCorner**，则此参数就以磅为单位指定从文档左上角到新线段终点的垂直距离。如果新线段的 *EditingType* 为 **msoEditingAuto**，则不用为此参数指定值。 |

## 示例

```javascript
/*此示例选择活动工作表的第三个形状，检查该形状是否是一个任意多边形对象，如果是，则插入一个节点。此示例假定在活动工作表中存在三个形状。*/
function test() {
    ActiveSheet.Shapes.Item(3).Select()
    if (Selection.ShapeRange.Type == msoFreeform) {
        Selection.ShapeRange.Nodes.Insert(3, msoSegmentCurve, msoEditingSymmetric, 35, 100)
        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(0, 0, 200)
        Selection.ShapeRange.Fill.Visible = msoTrue
    } else {
        console.log("This shape is not a Freeform object.")
    }
}
```

```javascript
/*本示例在第一张工作表中第二个形状末尾插入新的形状节点。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(2).Nodes
    nodes.Insert(nodes.Count + 1, msoSegmentLine, msoEditingAuto, 90, 100)
}
```
